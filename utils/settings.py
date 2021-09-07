import datetime as dt
from requests.structures import CaseInsensitiveDict
from .secrets_creation import *

## Retreive the API key from OS environment variables, create string variable for CS website name
try:
    with open(CS_CREDENTIALS_PATH) as f:
        data = json.load(f)
except FileNotFoundError:
    create_secrets()
    with open(CS_CREDENTIALS_PATH) as f:
        data = json.load(f)
CS_USER = data['CS_USER']
CS_PASS = data['CS_PASS']
API_KEY = data['CS_API']
CS = "http://live.makecontact.space/"
CS_API = "https://api.contactspace.com/?"
CS_THUNDER = "https://apithunder.makecontact.space/"
GT_API = 'https://givetel-api.com/v1/token'
headers = CaseInsensitiveDict()
headers['x-api-key'] = f"{API_KEY}"
CS_API_HEADERS = headers
HOME = os.getcwd()
RESOURCES_PATH = os.path.join(HOME, "Resources")
TEMP_PATH = os.path.join(HOME, "Temp")
LR_PATH = os.path.join(TEMP_PATH, "LRs")
ABI_PATH = os.path.join(TEMP_PATH, "Agent_By_Init")
IBA_PATH = os.path.join(TEMP_PATH, "Init_By_Agent")
RES_DAILY_PATH = os.path.join(TEMP_PATH, "Results_Daily")
RES_WEEKLY_PATH = os.path.join(TEMP_PATH, "Results_Weekly")
DATA_EXP_PATH = os.path.join(TEMP_PATH, "Data_Exports")
FUND_TIMES_PATH = os.path.join(TEMP_PATH, "Fundraiser_Times")
NEW_VS_EXISTING_PATH = os.path.join(TEMP_PATH, "New_Vs_Existing")
RECORDINGS_PATH = os.path.join(TEMP_PATH, "Call Recordings")
RESULTS_OUTPUT_FOLDERS = [
    LR_PATH,
    ABI_PATH,
    IBA_PATH,
    RES_DAILY_PATH,
    RES_WEEKLY_PATH,
    NEW_VS_EXISTING_PATH,
    RECORDINGS_PATH,
]
CS_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
### Department Email Definition
DATA_EMAILS = ["data_managers@company.com"]
QA_EMAILS = ["qa_team@company.com"]
COACHING_EMAILS = ["coaching_team@company.com"]
CLIENT_SERVICES_EMAILS = ["client_services@company.com"]
CEO_EMAIL = ["ceo@company.com"]
HR_EMAILS = ["hr@company.com"]
ASSISTANT_EMAIL = ["ops_assistant@company.com"]
OPS_EMAILS = QA_EMAILS + COACHING_EMAILS + HR_EMAILS + ASSISTANT_EMAIL
MANAGEMENT_EMAILS = OPS_EMAILS + DATA_EMAILS + CLIENT_SERVICES_EMAILS + CEO_EMAIL
EXCELWRITER_ENGINE = "xlsxwriter"
## Create datetime constants
TODAY = dt.date.today()
YESTERDAY = TODAY - dt.timedelta(days=1)
MONDAY = TODAY - dt.timedelta(days=TODAY.weekday())
LAST_MONDAY = MONDAY - dt.timedelta(days=7)
FN_MONDAY = LAST_MONDAY - dt.timedelta(days=7)
NEXT_MONDAY = MONDAY + dt.timedelta(days=7)
NEXT_FN_MONDAY = NEXT_MONDAY + dt.timedelta(days=7)
FRIDAY = TODAY - dt.timedelta(days=TODAY.weekday()) + dt.timedelta(days=4)
LAST_FRIDAY = (
    TODAY - dt.timedelta(days=TODAY.weekday()) - dt.timedelta(days=3)
)
START_OF_MONTH = TODAY
if START_OF_MONTH.day > 25:
    START_OF_MONTH += dt.timedelta(7)
START_OF_MONTH = START_OF_MONTH.replace(day=1)