from .skillid import *
from .datetime import *
import requests

## Fetch a list of all users and overwrite "Givetel_Users.xlsx" in the resources folder
def download_gt_users():
    '''Get list of current GT users and their ID's.'''
    global ALL_USERS
    TARGET_COLUMNS = ["Id", "FirstName", "Surname", "Email", "CRM_ID"]
    records = []
    dfs = []
    url = f"{CS_API}apikey={API_KEY}&function=GetUsers&outputtype=json"
    r = requests.post(url)
    data = r.json()
    records.append(data["users"])
    for record in records:
        df = pd.DataFrame.from_records(record)
        dfs.append(df)
    df = pd.concat(dfs)
    df = pd.DataFrame(df, columns=TARGET_COLUMNS)
    df["Id"] = pd.to_numeric(df["Id"])
    df["Agent Name"] = df["FirstName"] + " " + df["Surname"]
    df.to_excel(
        os.path.join(RESOURCES_PATH, "Givetel_Users.xlsx"),
        sheet_name="Users",
        index=False,
    )
    ALL_USERS = pd.read_excel(os.path.join(RESOURCES_PATH, "Givetel_Users.xlsx"))

def get_user_id(full_name):
    '''Get user id from full name'''
    full_name = str(full_name)
    return ALL_USERS.loc[ALL_USERS["Agent Name"] == full_name, "Id"].item()

def get_user_name(user_id):
    '''Get user name from user id (pass as int)'''
    user_id = int(user_id)
    return ALL_USERS.loc[ALL_USERS["Id"] == user_id, "Agent Name"].item()

def get_user_ids_from_skill(skill_id):
    '''Get a list of user id's assigned to a skill.
    Needs to be updated to thunder API.'''
    skill_id = str(skill_id)
    records = []
    dfs = []
    user_ids_stripped = []
    url = "https://api.contactspace.com/?apikey={}&function=GetSkillsInfo&skillid={}&outputtype=json".format(
        API_KEY, skill_id
    )
    r = requests.post(url)
    data = r.json()
    try:
        records.append(data["skill" + str(skill_id)]["Agents"])
    except KeyError:
        return (
            user_ids_stripped  ## Return empty list if no agents are assigned to skill
        )
    for record in records:
        df = pd.DataFrame.from_records(record)
        dfs.append(df)
    df = pd.concat(dfs)
    user_ids = df.columns
    for user in user_ids:
        user_ids_stripped.append(int(user[4:]))
    return user_ids_stripped

def get_active_users():
    return DAILY_TEAMS["Agent Name"].values

def get_user_pay_rate(full_name):
    full_name = str(full_name)
    # Return the pay rate of a fundraiser using their full name
    return ALL_USERS.loc[ALL_USERS["Agent Name"] == full_name, "Agent Rate"].item()

@error_logger
def remove_user_from_skill(skill_id, user_id):
    '''Remove one specific user from a skill. Needs to be updated to Thunder API.'''
    url = "https://api.contactspace.com/?apikey={}&function=RemoveUserFromSkill&skillid={}&userid={}&outputtype=json".format(
        API_KEY, skill_id, user_id
    )
    r = requests.post(url)
    if "not found" in r.text:
        raise ValueError
    print(r.text)

def remove_all_users_from_skill(skill_id):
    '''Remove all fundraisers from a given campaign - skill id must be provided'''
    for user in get_user_ids_from_skill(skill_id):
        remove_user_from_skill(skill_id, user)

def remove_all_users_from_all_active_skills():
    '''Remove all fundraisers from all active skills'''
    for skill in ACTIVE_DF["SkillId"]:
        remove_all_users_from_skill(skill)

def remove_all_users_from_all_skills():
    '''Remove all fundraisers from all skills'''
    for skill in GT_SKILL_ID["SkillId"]:
        remove_all_users_from_skill(skill)

def assign_user_to_skill(user_id, skill_id):
    '''Assign user to a skill. Needs to be updated to use Thunder API.'''
    url = "https://api.contactspace.com/?apikey={}&function=AddEditUserToSkill&skillid={}&userid={}&skilllevel=10&outputtype=xml".format(
        API_KEY, skill_id, user_id
    )
    if requests.post(url).status_code == 200: print(f"Assigning user {str(user_id)} {get_user_name(user_id)} to skill {str(skill_id)}: {get_init_name_from_skill_id(skill_id)}... Success.")
    else: print(f"Assigning user {str(user_id)} {get_user_name(user_id)} to skill {str(skill_id)}: {get_init_name_from_skill_id(skill_id)}... Failed.")

def assign_teams():
    '''Take teams from daily teams df and assign them to the appropriate skills in CS.'''
    collated_teams_ids = pd.merge(DAILY_TEAMS, ACTIVE_DF, on="Initiative")
    for user, skill_id in zip(
        collated_teams_ids["Agent Name"], collated_teams_ids["SkillId"]
    ):
        assign_user_to_skill(get_user_id(user), skill_id)

def is_active_user(user):
    return user in get_active_users()