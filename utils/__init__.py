from .settings import *
from .Google_test import Create_Service
import pandas as pd
import numpy as np
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests

### Create results output folders if they don't exist
[os.makedirs(folder) for folder in RESULTS_OUTPUT_FOLDERS if not os.path.exists(folder)]
## Instantiate reference dataframes for skillID's, Outcome/OG relationships, header transformations and user list
GT_SKILL_ID = pd.read_excel(os.path.join(RESOURCES_PATH, "GiveTel_skillid.xlsx"))
OUTCOME_OG = pd.read_excel(os.path.join(RESOURCES_PATH, "outcome_groups.xlsx"))
HEADERS = pd.read_excel(os.path.join(RESOURCES_PATH, "headers.xlsx"))
ALL_USERS = pd.read_excel(os.path.join(RESOURCES_PATH, "Givetel_Users.xlsx"))
DAILY_TEAMS = pd.read_excel(os.path.join(RESOURCES_PATH, "Daily_Teams.xlsx"))
try:
    CLI_SHEET = pd.read_excel(os.path.join(RESOURCES_PATH, "Givetel_CLIs.xlsx"))
except FileNotFoundError:
    print("No current CLI Sheet.")
DAILY_TEAMS = DAILY_TEAMS[["Agent Name", "Initiative"]]
# Drop row by index for each index where the corresponding team member isn't in the overall team list
[DAILY_TEAMS.drop(index=i, inplace=True) for i, name in enumerate(DAILY_TEAMS["Agent Name"]) if name not in ALL_USERS["Agent Name"].values]  
TARGET_HEADERS = HEADERS.columns.values.tolist()
ACTIVE_DF = GT_SKILL_ID[GT_SKILL_ID.Active != False]
if "CLI_SHEET" in locals():
    CLI_SHEET = CLI_SHEET[CLI_SHEET["Valid"] == True]
### Create dictionary of initiatives with their oldest open month for calling
OLDEST = {}
ACTIVE_DF = ACTIVE_DF.dropna(subset=["OLDEST_OPEN"])
[OLDEST.update({str(init): int(old)}) for init, old in zip(ACTIVE_DF["Initiative"], ACTIVE_DF["OLDEST_OPEN"])]

def ret_this(anything):
    # For lambda functions in pandas
    return anything

def to_bool(string):
    if str(string).lower() in ("n/a", "nan"):
        return np.nan
    if str(string).lower() in ("yes", "y", "true", "t", "1"):
        return True
    else:
        return False

def find_key(input_dict, value):
    return next((k for k, v in input_dict.items() if v == value), None)

def float_equals(f1, f2):
    return f1 == f2

def sum_numerical_dicts(dict_1, dict_2):
    '''
    Input: Two dictionaries with identical keys.
    Output: One dictionary with the keys and the values of both dictionaries summed.
    '''
    if dict_1.keys() != dict_2.keys(): raise ValueError("Dicts must have equal lengths and keys.")
    for key in dict_1: dict_1[key] += dict_2[key]
    return dict_1

### Decorators
def function_timer(f):
    def timer(*args, **kwargs):
        start = time.time()
        print(f"Timing {f.__name__}")
        rv = f(*args, **kwargs)
        end = time.time()
        print(f"{f.__name__} took {round((end-start), 6)} seconds to complete.")
        return rv
    return timer

def error_logger(f):
    def logger(*args, **kwargs):
        ### 
        # {Time: %(asctime)s}
        # {Function Name: %(funcName)s}
        # {Module: %(module)s}
        # {Level Name: %(levelname)s} - DEBUG, INFO, WARNING, ERROR, CRITICAL
        # {Line No: %(lineno)d} - Line of the code where log is issued
        # {message: %(message)s} - The logged message
        ###
        ERROR_LOG_FORMAT = "%(asctime)s:%(levelname)s:%(lineno)d:%(message)s"
        formatter = logging.Formatter(ERROR_LOG_FORMAT)
        logger = logging.getLogger(f.__name__)
        logger.setLevel(logging.ERROR)
        file_handler = logging.FileHandler(f"Logs/{f.__name__}.log")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        try:
            rv = f(*args, **kwargs)
        except ValueError:
            msg = f"{f.__name__} received an invalid value.\nGiven values were {[arg for arg in args]} and {kwargs.items()}."
            logger.exception(msg)
        except RuntimeError:
            msg = f"{f.__name__} had an unexperienced error, this could be due to a failed API call. Passed args were {[arg for arg in args]} and passed kwargs were {kwargs.items()}."
        except TypeError:
            logger.exception(f"{f.__name__} received and argument of an invalid type, or received the wrong amount of args.\nGiven args/types were {[(i, arg, 'Type: ' + str(type(arg))) for i, arg in enumerate(args)]} and {kwargs}")
        else:
            return rv
    return logger

def multi_threader(f): ### NOT WORKING
    def threader(*args, **kwargs):
        print(f.__name__)
        print(args, kwargs)
        threads = []
        result = []
        with ThreadPoolExecutor(max_workers=20) as executor:
            for item in args[0]:
                threads.append(executor.submit(f, [arg for arg in args])) ### f needs to be a second function that is passed to the decorator rather than the primary function
            for task in as_completed(threads):
                result.append(task.result())
            # [result.append(task.result()) for task in as_completed(threads)]
        return result
    return threader