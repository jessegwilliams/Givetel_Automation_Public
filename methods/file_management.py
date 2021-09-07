from utils import *
from shutil import copy

### Clear a folder of all its contents
def clear_folder(folder_path):
    for item in os.listdir(folder_path):
        try:
            os.remove(os.path.join(folder_path, item))
        except PermissionError:
            print('{} was not removed because the item is in use.'.format(item))

def delete_empty(folder_path):
    for f_name in os.listdir(folder_path):
        if "Results" in f_name:
            df = pd.read_excel(os.path.join(folder_path, f_name))
            columns = dict(map(reversed, enumerate(df.columns)))
            df.rename(columns=columns, inplace=True)
            calls_made = df.at[5, 2]
            if calls_made == 0:
                if f_name.startswith("All Agents"):
                    st_index = f_name.find("-")
                    en_index = f_name.find("Results")
                    substring = f_name[st_index + 2 : en_index - 1]
                    [
                        os.remove(os.path.join(folder_path, filename))
                        for filename in os.listdir(folder_path)
                        if substring in filename
                    ]
                else:
                    st_index = 0
                    en_index = f_name.find("-")
                    substring = f_name[st_index : en_index - 1]
                    [
                        os.remove(os.path.join(folder_path, filename))
                        for filename in os.listdir(folder_path)
                        if substring in filename
                    ]

def clear_raw_files(folder_path):
    ### Clear raw data files at the end of agent by init, init by agent, daily results or weekly results compilation
    for file_name in os.listdir(folder_path):
        if "All" in file_name or "download" in file_name:
            os.remove(os.path.join(folder_path, file_name))