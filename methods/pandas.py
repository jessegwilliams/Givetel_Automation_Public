if __name__ == '__main__':
    from outcomes import *
    from strings import *
    from skillid import *
    from file_management import *
else:
    from .outcomes import *
    from .strings import *
    from .skillid import *
    from .file_management import *
    from .datetime import date_to_gt_format
import xlrd

def save_to_excel(df, path, date_format='ddmmyy'):
    '''Function to save any dataframe or group of dataframes to an excel workbook.
    Path should be a full path including file name and extension.
    df can be a single df or a dictionary of dfs with sheetnames as keys and dataframes as values.'''
    with pd.ExcelWriter(path, date_format='ddmmyy') as writer:
        if type(df) == dict:
            for sheet_name, frame in df.items():
                if not frame is None:
                    frame.to_excel(writer, sheet_name, float_format="%.2f")
                    worksheet = writer.sheets[sheet_name]
                    worksheet = resize_columns(frame, worksheet)
                    writer.save
        elif type(df) == pd.DataFrame:
            sheet_name = 'Sheet 1'
            df.to_excel(writer, sheet_name, float_format='%.2f')
            worksheet = writer.sheets[sheet_name]
            worksheet = resize_columns(df, worksheet)
            writer.save

def resize_columns(df, worksheet):
    '''Iterate through columns of an xlsxwriter worksheet and make the columns the width of their widest member.'''
    for idx, col in enumerate(df.columns):  # loop through all columns
        series = df[col]
        max_len = (
            max(
                (
                    series.astype(str)
                    .map(len)
                    .max(),  # len of largest item
                    len(str(series.name)),  # len of column name/header
                )
            )
            + 1
        )  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column wid
    return worksheet

def cut_four_rows(df):
    '''Pandas utility to cut the top for rows of raw contactspace data and replace the column names with values from the first row'''
    df.drop(df.index[0:3], axis=0, inplace=True)
    df.columns = df.iloc[3]
    df.reset_index(inplace=True)
    df.drop(columns=["index"], axis=1, inplace=True)
    df.columns = df.iloc[0]
    df.drop(df.index[0], inplace=True)
    df.reset_index(inplace=True)
    df.drop(df.tail(2).index, inplace=True)
    df.reset_index(inplace=True)
    df.drop(columns=["index"], axis=1, inplace=True)
    df.drop(columns=["level_0"], axis=1, inplace=True)
    return df

def pivot_results(df, initiative_name, agent_name):
    necessary_cols = [
        "Answer Machine",
        "Call Back - Agent",
        "Call Back - Queue",
        "Chargeable",
        "IR & NS",
        "Regular Gifts",
        "Unchargeable",
    ]
    df.dropna(axis=1, how="all", inplace=True)
    for col in necessary_cols:
        if col not in df.columns:
            df[col] = 0
    try:
        df["Outcome_Group"] = df.apply(lambda row: outcome_to_og(row.Outcome), axis=1)
    except AttributeError:
        df["Outcome_Group"] = df.apply(lambda row: ret_this(0), axis=1)
    df["Contact"] = df.apply(lambda row: is_contact(row.Outcome), axis=1)
    df["DoubleGift"] = df.apply(lambda row: is_double_gift(row.Outcome), axis=1)
    results_table = pd.pivot_table(
        df,
        values="Count",
        index=df.columns[0],
        columns="Outcome_Group",
        aggfunc=np.sum,
        fill_value=0,
    )
    for col in necessary_cols:
        if col not in results_table.columns:
            results_table[col] = 0
    if "Double Gift" in results_table.columns:
        results_table["Contacts"] = (
            results_table["Chargeable"]
            + results_table["IR & NS"]
            + results_table["Regular Gifts"]
            + results_table["Double Gift"]
        )
        results_table["Regular Gifts"] = results_table["Regular Gifts"] + (
            2 * results_table["Double Gift"]
        )
    else:
        results_table["Contacts"] = (
            results_table["Chargeable"]
            + results_table["IR & NS"]
            + results_table["Regular Gifts"]
        )
    results_table.reset_index(inplace=True)
    if "Initiative Name" in results_table.columns:
        results_table["Agent Name"] = agent_name
    return results_table

def prepare_results_folder(folder_path):
    dfs = []
    last_index = folder_path.rfind("\\")
    overall_type = folder_path[last_index + 1 :]
    for spreadsheet in os.listdir(folder_path):
        f_name = os.path.join(folder_path, spreadsheet)
        df = pd.read_excel(f_name)
        columns = dict(map(reversed, enumerate(df.columns)))
        df.rename(columns=columns, inplace=True)
        agent_name = df.at[1, 4]
        initiative_name = df.at[1, 3]
        report_label = df.at[3, 0]
        df = cut_four_rows(df)
        if "Results" in spreadsheet:
            sheet_type = "Results"
            df = pivot_results(
                df,
                initiative_name=initiative_name,
                agent_name=agent_name,
            )
        elif "Hours" in spreadsheet:
            sheet_type = "Hours"
            df.drop(df.columns[[1, 2, 3, 4, 5]], axis=1, inplace=True)
        elif "Wrap" in spreadsheet:
            sheet_type = "Wrap"
            df.drop(df.columns[[1, 2, 3, 4, 5, 6, 7, 8, 11, 12]], axis=1, inplace=True)
            df["Wrap Time %"] = df["Wrap Time %"].str.rstrip("%").astype("float") / 100
            df["Pause Time %"] = (
                df["Pause Time %"].str.rstrip("%").astype("float") / 100
            )
            # print(df['Pause Time %'])
        elif "Average" in spreadsheet:
            sheet_type = "Average"
            df.drop(df.columns[[1, 2, 3, 4, 5, 6, 7, 8, 9, 10]], axis=1, inplace=True)
            df["Average Amount"] = df["Average Amount"].str.strip("$").astype("float")
            # print(df['Average Amount'])
        if agent_name == "All Agents" and initiative_name == "All Initiatives":
            df["Name"] = df.apply(lambda row: ret_this(report_label), axis=1)
        else:
            if agent_name != "All Agents":
                df["Name"] = df.apply(lambda row: ret_this(agent_name), axis=1)
            elif initiative_name != "All Initiatives":
                df["Name"] = df.apply(lambda row: ret_this(initiative_name), axis=1)
        df["Type"] = df.apply(lambda row: ret_this(sheet_type), axis=1)
        dfs.append(df)
    return (dfs, overall_type)

def prepare_results_singular(df):
    # if 'Unnamed: 1' in df.columns:
    #     return None
    columns = dict(map(reversed, enumerate(df.columns)))
    df.rename(columns=columns, inplace=True)
    agent_name = df.at[1, 4]
    initiative_name = df.at[1, 3]
    report_label = df.at[3, 0]
    from_date = df.at[1, 1]
    to_date = df.at[1, 2]
    df = cut_four_rows(df)
    df = pivot_results(df, initiative_name=initiative_name, agent_name=agent_name)
    if agent_name == "All Agents" and initiative_name == "All Initiatives":
        df["Name"] = df.apply(lambda row: ret_this(report_label), axis=1)
    else:
        if agent_name != "All Agents":
            df["Name"] = df.apply(lambda row: ret_this(agent_name), axis=1)
        elif initiative_name != "All Initiatives":
            df["Name"] = df.apply(lambda row: ret_this(initiative_name), axis=1)
    print('Dataframe at the end of prepare results singular: \n')
    print(df)
    print('Dataframe columns: \n')
    print(df.columns)
    return df, (from_date, to_date)

def rename_results(folder_path):
    '''Rename all raw data files in a folder for future results compilation'''
    for excel_file in os.listdir(folder_path):
        # Load dataframes of each excel file
        f_name = os.path.join(folder_path, excel_file)
        df = pd.read_excel(f_name)
        columns = dict(map(reversed, enumerate(df.columns)))
        df.rename(columns=columns, inplace=True)
        # Store name variables for later use
        agent_name = df.at[1, 4]
        initiative_name = df.at[1, 3]
        if agent_name == "All Agents" and initiative_name == "All Initiatives":
            if df.at[3, 0] == "Initiative Name":
                label = "All Initiatives"
            else:
                label = "All Agents"
        else:
            label = "{} - {}".format(agent_name, initiative_name)
        # Store type & date variables for later use
        rep_type = df.at[1, 0]
        from_date = df.at[1, 1]
        to_date = df.at[1, 2]
        if rep_type == "Agent Status":
            rep_type = "Wrap"
        elif rep_type == "Agent Calling Activity by Outcome":
            rep_type = "Results"
        elif rep_type == "Agent Calling Activity - Logged in hours":
            rep_type = "Hours"
        elif rep_type == "Conversion Report":
            rep_type = "Average"
        # Rename files with accrued information
        name_str = "{} {} {} - {}.xlsx".format(label, rep_type, from_date, to_date)
        print("Renaming to {}".format(os.path.join(folder_path, name_str)))
        # Attempt to save to folder or alert if file already exists in folder.
        try:
            os.rename(src=f_name, dst=os.path.join(folder_path, name_str))
            print("Rename successful.")
        except FileExistsError:
            print("Rename failed - file already exists. Skipping {}".format(name_str))

### Take a list of results dfs and return a list of indexes for later use
def get_merge_indexes(dfs):
    dfs_dict = {}
    results_dfs_indexes = []
    hours_dfs_indexes = []
    wrap_dfs_indexes = []
    average_dfs_indexes = []
    # Create df_dict, which categorizes all dfs in a dictionary structured like {index: {Name: Type}}
    for i, df in enumerate(dfs):
        subject = pd.unique(df["Name"])
        subject = subject[0]
        rep_type = pd.unique(df["Type"])
        rep_type = rep_type[0]
        dfs_dict.update({i: {subject: rep_type}})
    # Arrange dfs_dict info into 4 lists - results, hours, wrap and average. Dictionary values of {index: Agent} are stored in these lists.
    for k, v in dfs_dict.items():
        for key, val in v.items():
            if val == "Results": results_dfs_indexes.append({key: k})
            elif val == "Hours": hours_dfs_indexes.append({key: k})
            elif val == "Wrap": wrap_dfs_indexes.append({key: k})
            elif val == "Average": average_dfs_indexes.append({key: k})
    all_indexes = [results_dfs_indexes, hours_dfs_indexes, wrap_dfs_indexes, average_dfs_indexes]
    return all_indexes

def compile_single_report(
    ### Maybe add a to_date parameter to compile_single_report.
    results,
    hours=None,
    wrap=None,
    average=None,
    from_date=None,
    init_name=None,
):
    if hours is not None and wrap is not None and average is not None:
        rep = "Res"
    else:
        rep = "LR"
    # Create total calls column
    results["Total Calls"] = (
        results["Answer Machine"]
        + results["Call Back - Agent"]
        + results["Call Back - Queue"]
        + results["Contacts"]
        + results["Unchargeable"]
    )
    # Compile results report in the case of normal results
    if rep == "Res":
        hours_cols = hours.columns.difference(results.columns).tolist()
        hours_cols.append(results.columns[0])
        wrap_cols = wrap.columns.difference(results.columns).tolist()
        wrap_cols.append(results.columns[0])
        average_cols = average.columns.difference(results.columns).tolist()
        average_cols.append(results.columns[0])
        # Remove test campaigns
        results = results[~results[results.columns[0]].str.contains("test")]
        results = results[~results[results.columns[0]].str.startswith("CS")]
        results = results[~results[results.columns[0]].str.startswith("lora")]
        results = results[~results[results.columns[0]].str.startswith("Lora")]
        # Add Total Row
        results.loc["Total"] = results.sum(numeric_only=True)
        # Merge all dataframes
        results = results.merge(hours[hours_cols], how="left", on=results.columns[0])
        results = results.merge(wrap[wrap_cols], how="left", on=results.columns[0])
        results = results.merge(
            average[average_cols], how="left", on=results.columns[0]
        )
        # Calculate total hours
        results.loc[len(results) - 1, "Logged In Time (Hours)"] = results[
            "Logged In Time (Hours)"
        ].sum(skipna=True)
        # Calculate CPH column
        try:
            results["CPH"] = (
                results["Contacts"]
                .divide(results["Logged In Time (Hours)"])
                .replace(np.inf, 0)
            )
        except ZeroDivisionError:
            results["CPH"] = 0
        # Create KPI Column
        if results.columns[0] == "Agent Name":
            if "Initiative Name" in results.columns:
                results.rename(
                    columns={"Initiative Name": "InitiativeName"}, inplace=True
                )
                results["KPI"] = results.apply(
                    lambda row: get_kpi(row.InitiativeName), axis=1
                )
                results.rename(
                    columns={"InitiativeName": "Initiative Name"}, inplace=True
                )
            else:
                if init_name is not None:
                    results["KPI"] = results.apply(
                        lambda row: get_kpi(init_name), axis=1
                    )
                else:
                    results = results.assign(KPI=1)
        elif results.columns[0] == "Initiative Name":
            results.rename(columns={"Initiative Name": "InitiativeName"}, inplace=True)
            results["KPI"] = results.apply(
                lambda row: get_kpi(row.InitiativeName), axis=1
            )
            results.rename(columns={"InitiativeName": "Initiative Name"}, inplace=True)
        else:
            results = results.assign(KPI=get_kpi(results.columns[0]))
    # Compile list report results in the case of list reports
    else:
        # Get rid of removed data
        results = results[~results[results.columns[0]].str.startswith("REMOVED")]
        results = results[~results[results.columns[0]].str.startswith("xDNC_returned")]
        results = results[~results[results.columns[0]].str.startswith("DNC_returned")]
        # Set KPI's for conversion and contact rate
        if len(results['Name']) < 1:
            return None
        cr = get_cr(results["Name"][0])
        cr_label = str(cr) + "%"
        cpc = get_cpc(results["Name"][0])
        kpi = get_kpi(results["Name"][0])
        # Create List number column
        results["List #"] = results[results.columns[0]].str.extract(
            pat=r"(\d+)", expand=True
        )
        results["List #"] = pd.to_numeric(results["List #"])
        # Create load date column
        results["LD"] = results[results.columns[0]].str.extract(
            pat=r"(\d{6})", expand=True
        )
        results["LD"] = pd.to_datetime(results["LD"], format=r"%d%m%y", dayfirst=True)
        # Remove lists loaded before the beginning of the export window
        index_names = results[(results["LD"] < from_date)].index
        results.drop(index_names, inplace=True)
        # Create columns for load month & year as well as lead count.
        results["LM"] = results["LD"].dt.month
        results["LY"] = results["LD"].dt.year
        results["Leads"] = results[results.columns[0]].str.extract(
            pat=r".*\((\d*)\).*", expand=True
        )
        results["Leads"] = pd.to_numeric(results["Leads"])
        # Create total row
        if len(results) > 0:
            results.loc["Total"] = results.sum(numeric_only=True)
        # Create KPI column
        results["KPI"] = results.apply(lambda row: ret_this(kpi), axis=1)
    # Create columns for percentages
    results["IR%"] = results["IR & NS"] / results["Contacts"]
    results["RG%"] = results["Regular Gifts"] / results["Contacts"]
    results["UC%"] = results["Unchargeable"] / (
        results["Unchargeable"] + results["Contacts"]
    )
    results["AM%"] = results["Answer Machine"] / results["Total Calls"]
    results["CB%"] = (
        results["Call Back - Agent"] + results["Call Back - Queue"]
    ) / results["Total Calls"]
    results["CBQ%"] = results["Call Back - Queue"] / results["Total Calls"]
    results["CBA%"] = results["Call Back - Agent"] / results["Total Calls"]
    results["Contact%"] = results["Contacts"] / results["Total Calls"]
    # Fill NA cells with zeroes
    results.fillna(0, inplace=True)
    # Label total row
    results[results.columns[0]].replace(0, "Total", inplace=True)
    # Create column denoting whether the row is over conversion target
    results["Over"] = results["RG%"] >= results["KPI"]
    # Order columns appropriately
    if rep == "Res":
        results = results[[results.columns[0], "Contacts", "IR & NS", "Regular Gifts", "RG%", "Unchargeable",
                "UC%", "Logged In Time (Hours)", "CPH", "IR%", "Wrap Time %", "Pause Time %", "Average Amount", "AM%", "CB%", "CBQ%", "CBA%", "Contact%", "KPI", "Over"]]
    else:
        # Create columns for [cr_kpi, 'True CR', 'CR', 'CTG', 'True CTG', 'Income']
        results[cr_label] = results["Leads"] * (cr)
        results["CR"] = results["Contacts"] / results["Leads"]
        results["True CR"] = results["Contacts"] / (
            results["Leads"] - results["Unchargeable"]
        )
        results["CTG"] = results["Contacts"] - results[cr_label]
        results["True CTG"] = results["Contacts"] - (
            results[cr_label] - results["Unchargeable"]
        )
        results["Income"] = results["Contacts"] * cpc
        # Create Monthly CTG column
        group = results.groupby("LM")
        l = pd.DataFrame.last_valid_index
        if len(group.CTG.sum().values) > 0:
            results.loc[group.apply(l), "Monthly CTG"] = group["CTG"].sum().values
        results.fillna("", inplace=True)
        try:
            results["Monthly CTG"] = pd.to_numeric(results["Monthly CTG"])
        except KeyError:
            results["Monthly CTG"] = results.apply(lambda row: ret_this(0), axis=1)
        # Create Peno column
        results["Peno"] = (results["Contacts"] + results["Unchargeable"]) / results[
            "Leads"
        ]
        # Empty total values from irrelvant fields
        if len(results) > 1:
            results.iloc[-1, results.columns.get_loc("LM")] = ""
            results.iloc[-1, results.columns.get_loc("List #")] = ""
            results.iloc[-1, results.columns.get_loc("LD")] = ""
            results.iloc[-1, results.columns.get_loc("LY")] = ""
        # Reorder column
        results = results[[results.columns[0], "LD", "LM", "LY", "List #", "Leads", "Contacts", "CR", f"{cr_label}", "CTG", "Chargeable", "IR & NS", "Regular Gifts", "RG%", "Unchargeable", "UC%",
                "Total Calls", "IR%", "Peno", "True CR", "True CTG", "Monthly CTG", "AM%", "CB%", "CBQ%", "CBA%", "Contact%", "Name", "Income", "KPI", "Over",]]
    return results

def compile_results(dfs, merge_indexes):
    results_indexes = merge_indexes[0]
    results_dfs = []
    for index in results_indexes:
        for _, v in index.items():
            results = dfs[v]
            hours = dfs[v - 1]
            wrap = dfs[v + 1]
            average = dfs[v - 2]
            # print(results, hours, wrap, average)
            name = results["Name"][0]
            if name in GT_SKILL_ID["Initiative"].values:
                init_name = name
            else:
                init_name = None
            results = compile_single_report(
                results, hours, wrap, average, init_name=init_name
            )
            results_dfs.append({name: results})
    return results_dfs

def results_to_excel(dfs, folder_path, report_type):
    with pd.ExcelWriter(os.path.join(folder_path, (report_type + ".xlsx"))) as writer:
        for entry in dfs:
            for name, df in entry.items():
                if len(name) > 30:
                    name = replace_name(name)
                df.to_excel(writer, name, index=False)
                workbook = writer.book
                worksheet = writer.sheets[name]
                currency_format = workbook.add_format({"num_format": "$0.00"})
                float_format = workbook.add_format({"num_format": "0.00"})
                percentage_format = workbook.add_format({"num_format": "0.00%"})
                worksheet.set_column("E:E", None, percentage_format)
                worksheet.set_column("G:G", None, percentage_format)
                worksheet.set_column("J:R", None, percentage_format)
                worksheet.set_column("H:I", None, float_format)
                worksheet.set_column("M:M", None, currency_format)
                for idx, col in enumerate(df.columns):  # loop through all columns
                    series = df[col]
                    max_len = (
                        max(
                            (
                                series.astype(str)
                                .map(len)
                                .max(),  # len of largest item
                                len(str(series.name)),  # len of column name/header
                            )
                        )
                        + 1
                    )  # adding a little extra space
                    worksheet.set_column(idx, idx, max_len)  # set column wid
                writer.save

def list_rep_to_excel(df, folder_path):
    if not df is None:
        df.reset_index(inplace=True)
    else:
        return None
    if not len(df) == 0:
        name = df["Name"][0]
        df.drop(columns=["Name", "index"], inplace=True)
        with pd.ExcelWriter(
            os.path.join(folder_path, ("{} {}.xlsx".format(name, TODAY))),
            date_format="ddmmyy",
        ) as writer:
            if len(name) > 30:
                name = replace_name(name.upper())
            df.to_excel(writer, name, index=False)
            workbook = writer.book
            worksheet = writer.sheets[name]
            currency_format = workbook.add_format({"num_format": "$0.00"})
            float_format = workbook.add_format({"num_format": "0.00"})
            percentage_format = workbook.add_format({"num_format": "0.00%"})
            worksheet.set_column("H:H", None, percentage_format)
            worksheet.set_column("J:J", None, float_format)
            worksheet.set_column("N:N", None, percentage_format)
            worksheet.set_column("P:P", None, percentage_format)
            worksheet.set_column("R:R", None, percentage_format)
            worksheet.set_column("S:T", None, percentage_format)
            worksheet.set_column("W:AA", None, percentage_format)
            worksheet.set_column("AB:AB", None, currency_format)
            writer.save

def floorwide_ctgs():
    fwide_name = os.path.join(LR_PATH, "aaFloorwide CTG {}.xlsx".format(str(TODAY)))
    clear_raw_files(LR_PATH)
    files_updated = os.listdir(LR_PATH)
    # Empty lists for later use
    fnames = []
    ctgs = []
    dfs = []
    all_months = []
    monthly_ctgs = []
    # Define range of months
    for file_name in files_updated:
        whole_path = os.path.join(LR_PATH, file_name)
        df = pd.read_excel(whole_path)
        months = df["LM"].unique().tolist()
        months = [x for x in months if str(x) != "nan"]
        for month in months:
            if not re.search("[a-zA-Z]", str(month)):
                if not month in all_months:
                    all_months.append(month)
    ## Get CTG by month from all files and update list monthly_ctgs
    for file_name in files_updated:
        whole_path = os.path.join(LR_PATH, file_name)
        df = pd.read_excel(whole_path)
        # Create range of months to loop through
        months = df["LM"].unique().tolist()
        vals = {}
        for month in all_months:
            # Check if month is in df['LM'] and add {month: sum of CTG filtered by df['LM'][month]} to list monthly_ctgs
            if month in months:
                filt = df["LM"] == month
                vals.update({month: df.loc[filt]["CTG"].sum()})
            else:
                # If month not in df['LM'], add month: nan to dict.
                vals.update({month: np.nan})
        monthly_ctgs.append(vals)
    # Update ctg values for later complilation to fwide df
    for file_name in files_updated:
        whole_path = os.path.join(LR_PATH, file_name)
        df = pd.read_excel(whole_path)
        df.drop(df.columns[[0]], axis=1, inplace=True)
        val = df["CTG"].iloc[-1]
        if isinstance(val, (int, float)) == True:
            print("{} in {} is a number!".format(val, file_name))
            pass
        else:
            print("{} in {} is not a number".format(val, file_name))
            try:
                val = float(val)
            except:
                val = 0.0
        ctgs.append(val)
        dfs.append(df)
        fnames.append(replace_name(file_name))
    # Collate fwide ctg & monthly df lists into dataframes, then merge those dataframes to create the final
    d = {"Campaign": fnames, "CTG": ctgs}
    fwide_ctg_df = pd.DataFrame(d)
    monthly_df = pd.DataFrame(monthly_ctgs).fillna(0)
    collated_df = pd.merge(fwide_ctg_df, monthly_df, left_index=True, right_index=True)
    for col in collated_df.columns:
        if col != "Campaign" and col != "CTG":
            collated_df["%" + str(col)] = collated_df[col] / collated_df["CTG"]
    collated_df.loc["Total"] = collated_df.sum(numeric_only=True)
    for col in all_months:
        collated_df["%" + str(col)] = collated_df["%" + str(col)].map("{:.2%}".format)
    collated_df.to_excel(fwide_name, index=False)

def abi_front_sheet():
    file_path = os.path.join(ABI_PATH, "Agent_By_Init {}.xlsx".format(str(TODAY)))
    xlrd_book = xlrd.open_workbook(file_path, on_demand=True)
    tabs = xlrd_book.sheet_names()
    dfs = []
    camps_over_under = {}
    total_cont_over = {}
    total_cont_under = {}
    total_ty_over = {}
    total_ty_under = {}
    total_gifts = {}
    for tab in tabs:
        df = pd.read_excel(file_path, sheet_name=tab)
        df["Name"] = tab
        t_count = 0
        f_count = 0
        cont_over = 0
        cont_under = 0
        tys_over = 0
        tys_under = 0
        gifts = 0
        for i, value in enumerate(df["Over"].values):
            if (
                value
                and df["Initiative Name"][i] != "Total"
                and df["Logged In Time (Hours)"][i] > 0.0
            ):
                if (
                    "Thank You" not in df["Initiative Name"][i]
                    and "Welcome" not in df["Initiative Name"][i]
                    and "TY" not in df["Initiative Name"][i]
                ):
                    t_count += 1
                    cont_over += df["Contacts"][i]
                    gifts += df["Regular Gifts"][i]
                    # print('{} had {} gifts on {}, an {} target campaign'.format(tab, df['Regular Gifts'][i], df['Initiative Name'][i], value))
                else:
                    tys_over += df["Contacts"][i]
            elif (
                not value
                and df["Initiative Name"][i] != "Total"
                and df["Logged In Time (Hours)"][i] > 0.0
            ):
                if (
                    "TY" not in df["Initiative Name"][i]
                    and "Welcome" not in df["Initiative Name"][i]
                    and "Thank You" not in df["Initiative Name"][i]
                ):
                    f_count += 1
                    cont_under += df["Contacts"][i]
                    gifts += df["Regular Gifts"][i]
                else:
                    tys_under += df["Contacts"][i]
        camps_over_under.update({tab: [t_count, f_count]})
        total_cont_over.update({tab: cont_over})
        total_cont_under.update({tab: cont_under})
        total_ty_over.update({tab: tys_over})
        total_ty_under.update({tab: tys_under})
        total_gifts.update({tab: gifts})
        dfs.append(df)
    extra_columns = [
        total_cont_over,
        total_cont_under,
        total_gifts,
        total_ty_over,
        total_ty_under,
    ]
    column_names = ["Contacts +", "Contacts -", "Gifts", "Thank Yous +", "Thank Yous -"]
    front_sheet = pd.DataFrame.from_dict(
        camps_over_under, orient="index", columns=["Campaigns +", "Campaigns -"]
    )
    for column, column_name in zip(extra_columns, column_names):
        col_series = pd.Series(column)
        front_sheet.insert(
            len(front_sheet.columns), column_name, col_series, allow_duplicates=False
        )
    front_sheet = front_sheet.assign(
        Conversion_No_TY=front_sheet["Gifts"]
        / (front_sheet["Contacts +"] + front_sheet["Contacts -"])
    )
    summary_path = os.path.join(ABI_PATH, "Agent Results Summary {}.xlsx".format(TODAY))
    writer = pd.ExcelWriter(summary_path)
    front_sheet.to_excel(writer, "Summary")
    writer.save()
    writer.close()

def iba_front_sheet():
    file_path = os.path.join(IBA_PATH, "Init_By_Agent {}.xlsx".format(str(TODAY)))
    xlrd_book = xlrd.open_workbook(file_path, on_demand=True)
    tabs = xlrd_book.sheet_names()
    dfs = []
    agents_over_under = {}
    total_cont_over = {}
    total_cont_under = {}
    total_gifts = {}
    agents_over = {}
    agents_under = {}
    for tab in tabs:
        df = pd.read_excel(file_path, sheet_name=tab)
        df["Initiative"] = tab
        t_count = f_count = cont_over = cont_under = gifts = 0
        agent_over_list = []
        agent_under_list = []
        for i, value in enumerate(df["Over"].values):
            if (value and df["Agent Name"][i] != "Total" and df["Logged In Time (Hours)"][i] > 0.0):
                t_count += 1
                cont_over += df["Contacts"][i]
                gifts += df["Regular Gifts"][i]
                agent_over_list.append(df["Agent Name"][i])
            elif (value != True and df["Agent Name"][i] != "Total" and df["Logged In Time (Hours)"][i] > 0.0):
                f_count += 1
                cont_under += df["Contacts"][i]
                gifts += df["Regular Gifts"][i]
                agent_under_list.append(df["Agent Name"][i])
        agents_over_under.update({tab: [t_count, f_count]})
        total_cont_over.update({tab: cont_over})
        total_cont_under.update({tab: cont_under})
        total_gifts.update({tab: gifts})
        agents_over.update({tab: agent_over_list})
        agents_under.update({tab: agent_under_list})
        dfs.append(df)
    extra_columns = [total_cont_over, total_cont_under, total_gifts, agents_over, agents_under]
    column_names = ["Contacts +", "Contacts -", "Gifts", "Fundraisers Over", "Fundraisers Under"]
    front_sheet = pd.DataFrame.from_dict(agents_over_under, orient="index", columns=["Fundraisers +", "Fundraisers -"])
    for column, column_name in zip(extra_columns, column_names):
        col_series = pd.Series(column)
        front_sheet.insert(len(front_sheet.columns), column_name, col_series, allow_duplicates=False)
    summary_path = os.path.join(IBA_PATH, "Initiative Results Summary {}.xlsx".format(TODAY))
    writer = pd.ExcelWriter(summary_path)
    front_sheet.to_excel(writer, "Summary")
    writer.save()
    writer.close()

def separate_newies(df):
    inductions = ALL_USERS[['Agent Name', 'Induction Period (TRUE/FALSE)']]
    df = pd.merge(df, inductions, how='inner', on='Agent Name')
    df.drop(df.index[df['Logged In Time (Hours)'] == 0], inplace=True)
    df.replace(np.nan, False, inplace=True)
    grouped = df.groupby(df['Induction Period (TRUE/FALSE)'])
    try:
        newies = grouped.get_group(True)
    except KeyError:
        newies = None
    try:
        experienced = grouped.get_group(False)
    except KeyError:
        experienced = None
    if newies is not None and len(newies) > 0:
        newies.loc['Total'] = newies.sum(numeric_only=True)
        newies['CPH'] = newies['Contacts']/newies['Logged In Time (Hours)']
        newies['RG%'] = newies['Regular Gifts']/newies['Contacts']
        newies = newies[['Agent Name', 'Contacts', 'Regular Gifts', 'RG%', 'CPH', 'Average Amount']]
    if experienced is not None and len(experienced) > 0:
        experienced.loc['Total'] = experienced.sum(numeric_only=True)
        experienced['CPH'] = experienced['Contacts']/experienced['Logged In Time (Hours)']
        experienced['RG%'] = experienced['Regular Gifts']/experienced['Contacts']
        experienced = experienced[['Agent Name', 'Contacts', 'Regular Gifts', 'RG%', 'CPH', 'Average Amount']]
    if newies is None and experienced is None:
        return None
    else:
        df = pd.concat([experienced, newies], keys = ['Experienced', 'New'])
    return df

def separate_all_newies(source_date=None):
    if isinstance(source_date, dt.date):
        gt_source_date = date_to_gt_format(source_date)
        source_date = str(source_date)
    else: raise TypeError("Date parameter must be datetime.date")
    SOURCE_FOLDER = os.path.join(IBA_PATH, gt_source_date)
    TARGET_FOLDER = os.path.join(IBA_PATH, gt_source_date)
    TARGET_FOLDER = os.path.join(TARGET_FOLDER)
    TARGET_PATH = os.path.join(TARGET_FOLDER, f"New vs Existing {source_date}.xlsx")
    SOURCE_FILE = os.path.join(SOURCE_FOLDER, f'Init_By_Agent {source_date}.xlsx')
    xl = pd.ExcelFile(SOURCE_FILE)
    dfs = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(SOURCE_FILE, sheet)
        dfs.append({sheet: separate_newies(df)})
    if not len(dfs) == 0:
        with pd.ExcelWriter(TARGET_PATH, date_format="ddmmyy") as writer:
            for entry in dfs:
                for sheet_label, df in entry.items():
                    if not df is None:
                        df.to_excel(writer, sheet_label, float_format="%.2f")
                        workbook = writer.book
                        worksheet = writer.sheets[sheet_label]
                        int_format = workbook.add_format({"num_format": "0"})
                        currency_format = workbook.add_format({"num_format": "$0.00"})
                        float_format = workbook.add_format({"num_format": "0.00"})
                        percentage_format = workbook.add_format({"num_format": "0.00%"})
                        worksheet.set_column("D:E", None, int_format)
                        worksheet.set_column("F:F", None, percentage_format)
                        worksheet.set_column("G:G", None, float_format)
                        worksheet.set_column("H:H", None, currency_format)
                        worksheet = resize_columns(df, worksheet)
                        writer.save

def process_single_lr_file(folder, file_name):
    full_path = os.path.join(folder, file_name)
    df = pd.read_excel(full_path)
    if df["Unnamed: 2"].values[5] == 0:
        print("No results for {}".format(file_name))
    else:
        prep = prepare_results_singular(df)
        prepared = prep[0]
        from_date = prep[1][0]
        list_rep_to_excel(
            ### Maybe add a to_date parameter to compile_single_report.
            compile_single_report(prepared, from_date=from_date),
            folder,
        )