import re

### Replace long file names with short file names - pure string manipulation, no resources associated
def replace_name(name):
    name = str(name)
    ## Regex to remove version labelling from campaigns ie v2 from greenpeace petition v2
    name = re.sub(r"\sV\d\s", " ", name)
    ## Regex to remove list report label suffixes
    name = re.sub(r"LR\s\d\d\.\d\d\s\d\d\d\d\d\d\.xlsx", "", name)
    if "Agent Results" in name:
        name = name.replace("Agent Results", "")
    if "PETITION" in name:
        name = name.replace("PETITION", "PET")
    if "LEAD CONVERSION" in name:
        name = name.replace("LEAD CONVERSION", "LC")
    if "REACTIVATIONS" in name:
        name = name.replace("REACTIVATIONS", "REACT")
    if "UPGRADES" in name:
        name = name.replace("UPGRADES", "UG")
    if "PHONE SURVEY" in name:
        name = name.replace("PHONE SURVEY", "PS")
    if "WELCOME CALLS" in name:
        name = name.replace("WELCOME CALLS", "TY")
    if "VOTE2DONATE" in name:
        name = name.replace("VOTE2DONATE", "V2D")
    if "SHOW YOU CARE" in name:
        name = name.replace("SHOW YOU CARE", "SYC")
    if "NEW LAPSED" in name:
        name = name.replace("NEW LAPSED", "NL")
    if "REACTIVATION" in name:
        name = name.replace("REACTIVATION", "REACT")
    if "RETENTION" in name:
        name = name.replace("RETENTION", "RET")
    if "REJECTIONS" in name:
        name = name.replace("REJECTIONS", "REJ")
    if "BANK REJECTS" in name:
        name = name.replace("BANK REJECTS", "BR")
    if "CASH CONVERSION" in name:
        name = name.replace("CASH CONVERSION", "CASH")
    if "BEQUEST" in name:
        name = name.replace("BEQUEST", "BQ")
    if "THANK YOU" in name:
        name = name.replace("THANK YOU", "TY")
    if "RECYCLED" in name:
        name = name.replace("RECYCLED", "RC")
    if "BLACK DOG" in name:
        name = name.replace("BLACK DOG", "BDI")
    if "GREENPEACE" in name:
        name = name.replace("GREENPEACE", "GP")
    if "TARONGA" in name:
        name = name.replace("TARONGA", "TAR")
    if "THE SMITH FAMILY" in name:
        name = name.replace("THE SMITH FAMILY", "TSF")
    if "CC EXPIRY" in name:
        name = name.replace("CC EXPIRY", "CC EXP")
    if "TAX APPEAL" in name:
        name = name.replace("TAX APPEAL", "TAX")
    if "MVD" in name:
        name = name.replace("MVD", "TAX")
    if "JEWISH HOUSE CASH CONVERSION" in name:
        name = name.replace("JEWISH HOUSE CASH CONVERSION", "JH TAX")
    name = re.sub(r"\s\s+", "", name)
    return name