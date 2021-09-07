from utils import *

## Check a header name against a list of approved header names return the approved equivalent to the given header
def get_header_name(current_header):
    current_header = str(current_header)
    if current_header in TARGET_HEADERS:
        ## Return given value if it's already approved
        return current_header
    elif current_header not in HEADERS.values:
        ## Else check if the header isn't in the sheet
        print(current_header + " is not an identified header.")
        return current_header
    else:
        ## Otherwise find each instance of the value and return in list form
        find_value = HEADERS.isin([current_header])
        find_value_series = find_value.any()
        a = find_value_series[find_value_series == True].index.tolist()
        return a[0]

def transform_header(header):
    if header in HEADERS.columns:
        return header
    search = HEADERS.isin([header]).any()
    for i, item in enumerate(search):
        if item:
            return search.index[i]
    return header

def transform_all_headers(df):
    for column in df:
        df.rename(columns={column: transform_header(column)}, inplace=True)
    return df