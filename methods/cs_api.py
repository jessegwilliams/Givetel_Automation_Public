from json.decoder import JSONDecodeError
from utils import *
from .skillid import get_init_name
from .strings import replace_name
from .headers import transform_all_headers
from .users import *
from .file_management import *

def get_daily_call_ids(init_id, date=None):
    '''Returns all call ids for a given date (defaults to today). Useful for passing to get_record.'''
    from_date, to_date = date, date
    if from_date is None:
        from_date = str(TODAY)
    else:
        from_date = str(from_date)
    if to_date is None:
        to_date = str(TODAY)
    else:
        to_date = str(to_date)
    init_id = str(init_id)
    records = []
    dfs = []
    url = "{}apikey={}&function=GetCallIDs&fromdate={}&todate={}&initiativeid={}&outputtype=json&page=1".format(
        CS_API, API_KEY, from_date, to_date, init_id
    )
    r = requests.post(url)
    if r.text == "[]":
        return None
    data = r.json()
    records.append(data["records"])
    for record in records:
        df = pd.DataFrame.from_records(record)
        dfs.append(df)
    df = pd.concat(dfs)
    return df

def get_record(call_id=None):
    '''Returns the record associated with a certain call id. 
    Should be used to piece together datasets rather than using get_data. 
    Needs to be updated to use Thunder API.'''
    if call_id is None:
        return None
    else:
        call_id = str(call_id)
    records = []
    dfs = []
    url = "{}apikey={}&function=GetRecord&module=data&callid={}&outputtype=json".format(
        CS_API, API_KEY, call_id
    )
    r = requests.post(url)
    data = r.json()
    records.append(data["records"])
    for record in records:
        df = pd.DataFrame.from_records(record)
        dfs.append(df)
    df = pd.concat(dfs)
    df.rename(columns={"record": call_id}, inplace=True)
    df = df.transpose()
    return df

def get_daily_data(init_id, date=None, transform_headers=False, to_excel=False):
    '''Get data from the CS API for a given initiative ID. Defaults to today.
    Transform_headers uses Resources/headers.xlsx to make important header names uniform across campaigns. Newer campaigns may need their header names added to headers.xlsx.
    to_excel prints the data download to a spreadsheet.
    Needs to be updated to use the CS Thunder API. Threading should be removed in this update as CS throttles pull rates in the Thunder API.'''
    threads = []
    init_id = str(init_id)
    init_name = get_init_name(init_id)
    dfs = []
    # Get relevant call ID's
    print(f"Getting call ids for {init_name}...")
    call_ids_df = get_daily_call_ids(init_id, date)
    print("Done.")
    # Pull records from the CS API using the call ID's and compile into dfs list.
    if call_ids_df is not None:
        print(f"Getting records for call ids from {init_name}...")
        with ThreadPoolExecutor(max_workers=20) as executor:
            for call_id in call_ids_df["Id"]:
                threads.append(executor.submit(get_record, call_id))

            for task in as_completed(threads):
                dfs.append(task.result())
        print("Done.")
    else:
        return None
    data = dfs.pop(0)
    folder = os.path.join(DATA_EXP_PATH, str(TODAY))
    if not os.path.exists(folder):
        os.makedirs(folder)
    for df in dfs:
        data = data.append(df)
    # Transform headers if desired
    if transform_headers:
        data = transform_all_headers(data)
    file_name = get_init_name(init_id) + ".xlsx"
    # Print to excel
    if to_excel:
        data.to_excel(
            os.path.join(folder, file_name), sheet_name=replace_name(get_init_name(init_id))
        )
    return data

def get_fundraiser_times(user_id, date=None):
    # Status IDs: 0: Never Logged In. 1: Logged In. 2: Waiting. 3: Paused. 4: On Call. 5: Logged Out.
    # Status IDs: 6: System Assigning Record. 7: Previewing Record. 8: Preview Dialling Record. 9: Manual Dialling Record.
    # Status IDs: 10: Transferring Record. 11: In conference. 12: Wrapping.
    # Define list of each individual status ID.
    # Loop through each status id retrieving user's info for that status and collate all at the end.
    if date is None:
        date = str(TODAY)
    else:
        date = str(date)
    # Define range of statuses
    statuses = {0: "Never Logged In", 1: "Logged In", 2: "Waiting", 3: "Paused", 4: "On Call", 5: "Logged Out", 6: "System Assigning Record",
        7: "Previewing Record", 8: "Preview Dialling Record", 9: "Manual Dialling Record", 10: "Transferring Record", 11: "In conference", 12: "Wrapping"}
    all_times = []
    for status_id in statuses.keys():
        url = "{}apikey={}&function=AgentStatusHistory&userid={}&statusid={}&date={}&outputtype=json".format(CS_API, API_KEY, str(user_id), str(status_id), date)
        r = requests.post(url)
        data = r.json()
        if "ErrorMessage" in data:
            # Handle error for fundraisers with no times that week
            print("No {} time for {}.".format(statuses[status_id], get_user_name(user_id)))
        else:
            ## Add status id and status as dictionary entries for list of times
            [activity.update({"StatusID": status_id}) for activity in data["activities"]]
            [activity.update({"Status": statuses.get(status_id)}) for activity in data["activities"]]
            [all_times.append(activity) for activity in data["activities"]]
    if len(all_times) < 1:
        ## Return none if there are no times that day for a fundraiser
        return None
    times_df = pd.DataFrame(all_times)
    times_df["StartTime"] = pd.to_datetime(times_df["StartTime"], format=CS_DATE_FORMAT)
    times_df["EndTime"] = pd.to_datetime(times_df["EndTime"], format=CS_DATE_FORMAT)
    times_df["Duration"] = times_df["EndTime"] - times_df["StartTime"]
    times_df["Duration"] = times_df["Duration"] / np.timedelta64(1, "h")
    times_df = times_df.sort_values(by=["StartTime", "EndTime"])
    total = times_df["Duration"].sum()
    times_df.loc["Total"] = pd.Series(total, index=["Duration"])
    times_df = times_df[
        ["StatusID", "Status", "CallId", "StartTime", "EndTime", "Duration"]
    ]
    return times_df

@function_timer
def get_all_fundraiser_times():
    '''Accrue logs of all fundraiser times, collate to an excel workbook and optionally email to HR.'''
    dfs = []
    fundraisers = []
    target_path = os.path.join(
        FUND_TIMES_PATH, "Fundraiser Times {}.xlsx".format(str(TODAY))
    )
    for fundraiser in get_active_users():
        df = get_fundraiser_times(get_user_id(fundraiser))
        if df is not None:
            dfs.append(df)
            fundraisers.append(fundraiser)
    writer = pd.ExcelWriter(
        target_path,
        engine=EXCELWRITER_ENGINE,
    )
    for df, fundraiser in zip(dfs, fundraisers):
        df.to_excel(writer, sheet_name=fundraiser, index=False)
        worksheet = writer.sheets[fundraiser]
        for idx, col in enumerate(df):
            series = df[col]
            max_len = (
                max((series.astype(str).map(len).max(), len(str(series.name)))) + 1
            )
            worksheet.set_column(idx, idx, max_len)
    writer.save()

@error_logger
def switch_all_fundraisers(camp, new_camp):
    skill_id = get_skill_id(camp)
    new_skill_id = get_skill_id(new_camp)
    user_ids = get_user_ids_from_skill(skill_id)
    print(f"Current users calling {camp} are {[get_user_name(user_id) for user_id in user_ids]}")
    remove_all_users_from_skill(skill_id)
    print(f"Remaining users calling {camp} are {[get_user_name(user_id) for user_id in get_user_ids_from_skill(skill_id)]}")
    [assign_user_to_skill(user_id, new_skill_id) for user_id in user_ids]
    new_camp_user_ids = get_user_ids_from_skill(new_skill_id)
    print(f'Current users calling {new_camp} are {[user_id for user_id in new_camp_user_ids]}')

def get_recording(call_id, out_folder, outcome, agent_name, save, extra_fields_for_filename=None):
    '''Pull info about a call's recording from the CS Thunder API using the call id.'''
    if extra_fields_for_filename and not isinstance(extra_fields_for_filename, list): 
        extra_fields = list(extra_fields_for_filename)
    elif extra_fields_for_filename and isinstance(extra_fields_for_filename, list):
        extra_fields = extra_fields_for_filename
    else: extra_fields = []
    if not outcome is None:
        fn_elements = [outcome, agent_name]
        if len(extra_fields) > 0:
            for field in extra_fields:
                fn_elements.append(field)
        fn_end = " ".join(fn_elements)
    else: fn_end = ""
    headers = CS_API_HEADERS
    body = {'callid': call_id}
    url = f'{CS_THUNDER}GetRecording'
    r = requests.post(url, data=body, headers=headers)
    response_json = r.json()
    print(response_json)
    if 'error' in response_json.keys():
        print(f"{response_json['error']} for {call_id}")
        return None
    recording_url = response_json['recordings'][0]['URL']
    recording_date = response_json['recordings'][0]['Date']
    if not 'Recording not available through API.' in recording_url:
        recording_return = requests.get(recording_url)
        if save == True:
            with open(os.path.join(out_folder, f'{call_id} {fn_end} {recording_date}.wav'), 'wb') as f:
                f.write(recording_return.content)
    return r.text

def get_record_thunder(call_id):
    records = []
    dfs = []
    headers = CS_API_HEADERS
    body = {'callid': str(call_id), 'module': 'data'}
    url = "{}GetRecord".format(CS_THUNDER)
    r = requests.post(url, headers=headers, data=body)
    json_text = r.text
    ### Remove erroneous linebreaks from response before loading as json to assist decoding.
    if chr(10) in json_text: json.text = json_text.replace(chr(10), "//n")
    if chr(13) in json_text: json.text = json_text.replace(chr(13), "//n")
    if "(MISSING)" in json_text: json.text = json_text.replace("(MISSING)", "")
    data = json.loads(json_text)
    if not 'error' in data.keys():
        records.append(data["records"])
        for record in records:
            df = pd.DataFrame.from_records(record)
            dfs.append(df)
        df = pd.concat(dfs)
        return df
    else:
        print(data['error'])

def get_campaign_recordings(init_name, date=None, outcomes_to_exclude=None, extra_fields_for_filename=None):
    exclusion_list = list(outcomes_to_exclude)
    if extra_fields_for_filename: extra_fields = list(extra_fields_for_filename)
    else: extra_fields = None
    if not date: date = TODAY
    date = str(date)
    out_folder = os.path.join(RECORDINGS_PATH, init_name)
    if not os.path.exists(out_folder): os.mkdir(out_folder)
    if not os.path.exists(os.path.join(out_folder, date)): os.mkdir(os.path.join(out_folder, date))
    out_folder = os.path.join(out_folder, date)
    init_id = get_init_id(init_name)
    call_ids = get_daily_call_ids(init_id, date)
    print(call_ids)
    if call_ids is None:
        print(f"No calls made for {init_name} on {date}.")
        return None
    for call_id in call_ids['Id']:
        record = get_record_thunder(call_id)
        if not record is None:
            extra_fields_to_pass = []
            outcome = record['Outcome'].values[0]
            agent_name = record['AgentName'].values[0]
            if extra_fields: 
                for field in extra_fields: extra_fields_to_pass.append(record[field].values[0])
            if not outcome in exclusion_list:
                if len(extra_fields_to_pass) > 0: get_recording(call_id, out_folder, outcome, agent_name, True, extra_fields_to_pass)
                else: get_recording(call_id, out_folder, outcome, agent_name, True)
            else:
                print(f"Didn't retreive {call_id} as it was a {outcome}.")