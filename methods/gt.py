### This script contains easily callable methods useful for automating Contactspace using their API or Selenium ###
## Import relevant packages
from .skillid import *
from .strings import *
from .datetime import *
from .users import *
from .pandas import *
from .headers import *
from .outcomes import *
from .cs_api import *
from .file_management import *

### Retrieve CLI information from Google Sheets and user information from Contactspace.
def refresh_resources():
    download_gt_users()

def send_campaign_summary(client=None, date=None, exclude=None, use_weekly_results=False):
    REDUCED_COLS = ['Initiative Name', 'Contacts', 'Regular Gifts', 'RG%']
    if use_weekly_results: source_result_label = "Results_Weekly"
    else: source_result_label = "Results_Daily"
    TARGET_FOLDER = os.path.join(HOME, 'Temp', 'Client_Services')
    if use_weekly_results:
        SOURCE_FOLDER = RES_WEEKLY_PATH
    else:
        SOURCE_FOLDER = RES_DAILY_PATH
    ## Control types and set path parameters
    if date is None: date = TODAY
    if not isinstance(date, dt.date):
        print('Date must be in datetime.date format.')
        return None
    if exclude is None: exclude = []
    results_folder = os.path.join(SOURCE_FOLDER, date_to_gt_format(date))
    TARGET_FOLDER = os.path.join(TARGET_FOLDER, client)
    TARGET_PATH = os.path.join(TARGET_FOLDER, "{} Results Summary {}.xlsx".format(client, date))
    ## Check if folder exists for date in Daily_Results
    if not os.path.exists(results_folder):
        print('Please create daily results for {}'.format(str(date)))
        return None
    ## Load initiative daily results from date into pandas dataframe
    results = pd.read_excel(os.path.join(results_folder, f'{source_result_label} {str(date)}.xlsx'), 'Initiative Name')
    ## Drop any rows that do not contain the campaign name in the Initiative Column
    if client: results = results[results[results.columns[0]].str.contains(client)]
    ## Drop any rows that include any of the exclude terms
    if exclude:
        for term in exclude: results = results[~results[results.columns[0]].str.contains(term)]
    ## Drop any rows where there are no contacts
    results.drop(results.index[results['Contacts'] == 0], inplace=True)
    ## Recalculate totals row and percentage columns
    results.loc["Total"] = results.sum(numeric_only=True)
    results['RG%'] = results['Regular Gifts']/results['Contacts']
    results['IR%'] = results['IR & NS']/results['Contacts']
    results['CPH'] = results['Contacts']/results['Logged In Time (Hours)']
    ## Drop non specified columns
    results = results[REDUCED_COLS]
    ## Set column formats and save in Analysis for Client Services/Campaign subfolder
    results.reset_index(inplace=True)
    results.drop(results.columns[0], inplace=True, axis=1)
    results.at[len(results)-1, 'Initiative Name'] = 'Total'
    results.rename(columns={"Initiative Name": "Campaigns"}, inplace=True)
    results.rename(columns={"Regular Gifts": "RGs"}, inplace=True)
    results.rename(columns={"RG%": "Conv"}, inplace=True)
    if not os.path.exists(TARGET_FOLDER):
        os.mkdir(TARGET_FOLDER)
    if not len(results) == 0:
        with pd.ExcelWriter(TARGET_PATH, date_format="ddmmyy") as writer:
            if len(client) > 30:
                name = replace_name(client.upper())
            else: name = client.upper()
            results.to_excel(writer, name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[name]
            int_format = workbook.add_format({"num_format": "0"})
            percentage_format = workbook.add_format({"num_format": "0.00%"})
            worksheet.set_column("B:C", None, int_format)
            worksheet.set_column("D:D", None, percentage_format)
            worksheet = resize_columns(results, worksheet)
            writer.save
    return results

@error_logger
def get_daily_data_for_active_campaigns(date=None):
    if date is None:
        date = TODAY
    dfs = []
    for init_id in ACTIVE_DF['InitiativeId']: 
        dfs.append(get_daily_data(init_id, date=date))
    return dfs

@error_logger
def get_gift_counts(df):
    if df is None: return df
    results = {"Gifts": 0, "Upgrades": 0, "Reacts": 0, "Thank Yous": 0, "Bequests": 0, "Admin": 0, "Tax Appeal": 0}
    init_name = replace_name(get_init_name(df['InitiativeId'].sample()))
    # results = results[~results[results.columns[0]].str.contains("test")]
    df['OutcomeGroup'] = df['Outcome'].apply(outcome_to_og)
    if 'Regular Gifts' not in df['OutcomeGroup'].array:
        return {init_name: [None, results]}
    df = df[df['OutcomeGroup'] == 'Regular Gifts']
    [df.drop(col, axis=1, inplace=True) for col in df.columns if col not in HEADERS.columns] ### Drop any non universal columns
    if not ('TY' in init_name or 'UG' in init_name or 'REACT' in init_name or 'REJ' in init_name or 'BR' in init_name or 'DEC' in init_name):
        results['Gifts'] += len(df)
    if 'UG' in init_name:
        results['Upgrades'] += len(df)
    if 'REACT' in init_name:
        results['Reacts'] += len(df)
    if 'TY' in init_name:
        results['Thank Yous'] += len(df)
    if 'BQ' in init_name:
        results['Bequests'] += len(df)
    if 'DEC' in init_name or 'REJ' in init_name or 'BR' in init_name or 'CC EXP' in init_name:
        results['Admin'] += len(df)
    if 'TAX' in init_name:
        results['Tax Appeal'] += len(df)
    return {init_name: [df, results]}

@error_logger
def get_missing_gifts(df, init_name):
    if df is None: return {}
    if len(df) == 0: return {}
    result = {}
    if 'Amount' in df.columns:
        for idx, amount in enumerate(df['Amount'].array):
            if (amount is None or float(amount) == 0.0) and not ('TY' in init_name or 'REJ' in init_name or 'BR' in init_name or 'DEC' in init_name):
                result.update({df['Phone1'].array[idx]: df['AgentName'].array[idx]})
    return result

def get_daily_confirms(date=None, email=False, recipients=None, dfs=None, save_to_excel=False):
    if not recipients: recipients = QA_EMAILS + COACHING_EMAILS + ASSISTANT_EMAIL
    print(f"Starting confirms collation.")
    start = time.time()
    if date is None:
        date = TODAY
    if dfs is None: dfs = get_daily_data_for_active_campaigns(date)
    TARGET_FOLDER = DATA_EXP_PATH
    TARGET_PATH = os.path.join(TARGET_FOLDER, str(date), "Confirms {}.xlsx".format(str(date)))
    confirms = {}
    results = {"Gifts": 0, "Upgrades": 0, "Reacts": 0, "Thank Yous": 0, "Bequests": 0, "Admin": 0, "Tax Appeal": 0}
    all_missing_gifts = {}
    for df in dfs:
        if not df is None:
            gift_counts = get_gift_counts(df)
            init_name = list(gift_counts)[0]
            if gift_counts[init_name][0] is None:
                next
            results = sum_numerical_dicts(results, gift_counts[init_name][1])
            confirms.update({init_name: gift_counts[init_name][0]})
            missing_gifts = get_missing_gifts(gift_counts[init_name][0], init_name)
            if not missing_gifts is None:
                all_missing_gifts.update(missing_gifts)
        else:
            print('No dataframe to add to confirms collection.')
    if save_to_excel: save_to_excel(confirms, TARGET_PATH)
    end = time.time()
    print(f"Confirms collation took {end-start} seconds.")
    if email:
        missing_gifts_str = ""
        if len(list(all_missing_gifts.keys())) > 0:
            for number, agent in all_missing_gifts.items():
                missing_gifts_str += f"\n{number} - {agent}"
    return confirms

def collate_records_of_criteria(df, search_field, search_value, exact_match=True):
    '''Input: A dataframe, the name of a header in that dataframe to search within, the value to search for in that field.
    Output: A dataframe containing records that match the search criteria.
    Not yet working.'''
    try:
        if not search_field in df.columns: raise KeyError("search_field must be a header name in the dataframe.")
        print(df[search_field].values)
    except AttributeError:
        return pd.DataFrame()
    if not search_value in df[search_field].values: raise ValueError(f"{search_value} is not contained in {search_field}")
    df = df[df[search_field].str.contains(search_value)]
    return df