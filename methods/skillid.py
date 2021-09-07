from utils import *

### Small functions using the givetel skill_id sheet ###
## Retreive initiative ID using its name
def get_init_id(init_name):
    init_name = str(init_name).upper()
    return GT_SKILL_ID.loc[
        GT_SKILL_ID["Initiative"] == init_name, "InitiativeId"
    ].item()

## Retrieve initiative name using its ID
def get_init_name(init_id):
    init_id = int(init_id)
    return GT_SKILL_ID.loc[GT_SKILL_ID["InitiativeId"] == init_id, "Initiative"].item()

## Retrieve initiative name using its Skill_ID
def get_init_name_from_skill_id(skill_id):
    skill_id = int(skill_id)
    return GT_SKILL_ID.loc[GT_SKILL_ID["SkillId"] == skill_id, "Initiative"].item()

## Retreive skill id using initiative name
def get_skill_id(init_name):
    init_name = str(init_name)
    return GT_SKILL_ID.loc[GT_SKILL_ID["Initiative"] == init_name, "SkillId"].item()

## Retreive conversion kpi using initiative name
def get_kpi(init_name):
    init_name = str(init_name)
    try:
        ref = init_name.upper()
    except AttributeError:
        print("Couldn't get KPI for {}".format(init_name))
        return 0.0
    if ref in GT_SKILL_ID["Initiative"].values:
        returner = GT_SKILL_ID.loc[GT_SKILL_ID["Initiative"] == ref, "CONV_KPI"].item()
        print("{} KPI is {}!".format(ref, returner))
        if isinstance(returner, str):
            returner = returner.replace("%", "")
            return float(float(returner) / 100.0)
        else:
            return float(float(returner / 100.0))
    else:
        print("Couldn't get KPI for {}".format(ref))
        return 0.0

## Retreive contact rate kpi using initiative name
def get_cr(init_name):
    init_name = str(init_name)
    # if campaign is in skill id sheet return kpi for that campaign
    ref = init_name.upper()
    if ref in GT_SKILL_ID["Initiative"].values:
        returner = GT_SKILL_ID.loc[GT_SKILL_ID["Initiative"] == ref, "CR_KPI"].item()
        if isinstance(returner, str):
            returner = returner.replace("%", "")
            return float(returner)
        else:
            return returner
    else:
        return 0.0

## Retreive CPC using initiative name
def get_cpc(init_name):
    init_name = str(init_name)
    # if campaign is in skill id sheet return CPC for that campaign
    ref = init_name.upper()
    if ref in GT_SKILL_ID["Initiative"].values:
        returner = GT_SKILL_ID.loc[GT_SKILL_ID["Initiative"] == ref, "CPC"].item()
        if isinstance(returner, str):
            returner = returner.replace("$", "")
            return float(returner)
        else:
            return returner
    else:
        return 13.0

# Retreive oldest open month for a campaign (pass campaign name as string)
def get_oldest_open(init_name):
    # if campaign is in skill id sheet return CPC for that campaign
    ref = init_name.upper()
    if ref in ACTIVE_DF["Initiative"].values:
        returner = ACTIVE_DF.loc[ACTIVE_DF["Initiative"] == ref, "OLDEST_OPEN"].item()
        return returner
    else:
        return 13.0