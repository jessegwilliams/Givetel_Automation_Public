from utils import *
from .skillid import get_init_id, get_init_name

def outcome_to_og(outcome):
    '''Retrieve the outcome group of a given outcome. If the outcome is the name of a group, return the outcome.'''
    outcome = str(outcome)
    # If outcome is already an outcome group, return outcome.
    if outcome in OUTCOME_OG["Outcome_Group"].values:
        return outcome
    # Else return relevent outcome group if it exists
    elif outcome in OUTCOME_OG["Outcome"].values:
        return OUTCOME_OG.loc[OUTCOME_OG["Outcome"] == outcome, "Outcome_Group"].item()
    else:
        # Else return the input
        return outcome

def is_contact(outcome):
    outcome = str(outcome)
    # Else return relevent outcome group if it exists
    if outcome in OUTCOME_OG["Outcome"].values:
        return OUTCOME_OG.loc[OUTCOME_OG["Outcome"] == outcome, "Contact"].item()
    # If outcome is already an outcome group, return outcome.
    elif outcome in OUTCOME_OG["Outcome_Group"].values:
        return OUTCOME_OG.loc[OUTCOME_OG["Outcome_Group"] == outcome, "Contact"].sample()
    else:
        # Else return the input
        return np.nan

def is_double_gift(outcome):
    # Ensure correct typing
    outcome = str(outcome)
    # Else return relevent outcome group if it exists
    if outcome in OUTCOME_OG["Outcome"].values:
        return OUTCOME_OG.loc[OUTCOME_OG["Outcome"] == outcome, "DoubleGift"].item()
    # If outcome is already an outcome group, return outcome.
    elif outcome in OUTCOME_OG["Outcome_Group"].values:
        return OUTCOME_OG.loc[
            OUTCOME_OG["Outcome_Group"] == outcome, "DoubleGift"
        ].item()
    else:
        # Else return the input
        return np.nan

### Get currently used outcomes for a given initiative.
def get_outcomes(init_id):
    init_id = str(init_id)
    dfs = []
    records = []
    url = "{}apikey={}&function=GetOutcome&module=data&initiativeid={}&outputtype=json".format(
        CS_API, API_KEY, init_id
    )
    r = requests.post(url)
    data = r.json()
    records.append(data["outcomes"])
    for record in records:
        df = pd.DataFrame.from_records(record)
        dfs.append(df)
    df = pd.concat(dfs)
    return df

### Utility function to ensure all outcomes are in the outcome/outcome group relationships spreadsheet.
def check_for_missing_outcomes():
    def single_init_check(init_id):
        print("Checking outcomes for {}".format(get_init_name(init_id)))
        outcomes = get_outcomes(init_id)
        for outcome in outcomes["Name"]:
            if not outcome in OUTCOME_OG["Outcome"].values:
                print("{} is missing".format(outcome))

    for init in ACTIVE_DF["Initiative"]:
        single_init_check(get_init_id(init))