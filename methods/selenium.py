### This script contains easily callable methods useful for automating Contactspace using Selenium ###

## Import relevant packages
from .gt import *
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

class MyDriver:
    def __init__(self, folder=None, full_path=None):
        self.folder = folder
        self.full_path = full_path
        self.options = webdriver.ChromeOptions()
        self.prefs = {}
        self.prefs["profile.default_content_settings.popups"] = 0
        self.prefs["download.directory_upgrade"] = True
        if self.full_path is None and self.folder is None:
            print("No directory passed - setting to temp")
            self.prefs["download.default_directory"] = TEMP_PATH
        elif self.full_path is not None and self.folder is not None:
            print(
                "Too many values - pass a folder from the current directory or a full path. Setting to temp."
            )
            self.prefs["download.default_directory"] = TEMP_PATH
        elif self.full_path is None:
            self.prefs["download.default_directory"] = os.path.join(
                os.getcwd(), folder
            )
            print(
                "Setting download directory to {}...".format(
                    os.path.join(os.getcwd(), folder)
                )
            )
        elif folder is None:
            self.prefs["download.default_directory"] = full_path
            print("Setting download directory to {}...".format(full_path))
        self.options.add_experimental_option("prefs", self.prefs)
        self.driver = webdriver.Chrome(
            ChromeDriverManager().install(), options=self.options
        )

    # Login to ContactSpace
    def login(self):
        self.driver.get(CS)
        self.use_name("user", CS_USER)
        self.use_name("pwd", CS_PASS)
        self.use_xpath('//*[@id="login_form"]/button')

    # Utility clicks an element of a certain xpath
    def use_xpath(self, element, keys=None):
        elem = self.driver.find_element_by_xpath(element)
        if not keys:
            elem.click()
        else:
            elem.send_keys(keys)

    # Utility clicks an element of a certain name
    def use_name(self, name, keys=None):
        elem = self.driver.find_element_by_name(name)
        if not keys:
            elem.click()
        else:
            elem.send_keys(keys)

    # Wait until all chrome downloads are finished before continuing script
    def finish_all_downloads_chrome(self):
        if not self.driver.current_url.startswith("chrome://downloads"):
            self.driver.get("chrome://downloads/")
        return self.driver.execute_script(
            """
            return document.querySelector('downloads-manager')
            .shadowRoot.querySelector('#downloadsList')
            .items.filter(e => e.state === 'COMPLETE')
            .map(e => e.filePath || e.file_path || e.fileUrl || e.file_url);
            """
        )

    def quick_download(self, report_type_id, f_date, t_date, measurement, init_id=0, agent_id=0, init=None, agent=None):
        url = f"{CS}admin/reportsagent2_new.php?id={str(report_type_id)}&fromdate={str(f_date)}&todate={str(t_date)}&initiative={str(init_id)}&agent={str(agent_id)}&measurement={str(measurement)}"
        if init is not None:
            url += "&selinitiatives={}".format(init)
            print(init)
        if agent is not None:
            url += "&selagents={}".format(agent)
        self.driver.get(url)
        self.driver.implicitly_wait(1)
        self.use_xpath("/html/body/div[2]/div[2]/div/div/div[1]/h3/button[2]")
        self.driver.implicitly_wait(5)

    # Download raw list report data for all active campaigns
    def get_active_campaigns_lr_data(self):
        # change_download_directory(r"Temp\LRs")
        def get_lr_data(init_id, init, from_date):
            if from_date > TODAY:
                from_date = from_date - dt.relativedelta(years=1)
            self.quick_download(3, from_date, TODAY, 5, init_id, init=init)
        i = 0
        items = list(OLDEST.items())
        num_campaigns = len(items) - 1
        while i <= num_campaigns:
            get_lr_data(
                get_init_id(items[i][0]), items[i][0], first_of(items[i][1])
            )
            i += 1

    # Status ID for Wrap Time Values is 1, Results by Outcome is 3, Logged In Hours is 5
    def get_agents_by_init_for_week(self, last=False, inc_dolla=True):
        def get_result(agent_id, agent, last, id_id):
            measurement = 1
            if last:
                st_date = LAST_MONDAY
                en_date = LAST_FRIDAY
            else:
                st_date = MONDAY
                en_date = FRIDAY
            self.quick_download(
                id_id, st_date, en_date, measurement, agent_id=agent_id, agent=agent
            )

        last = last
        i = 0
        items = get_active_users()
        num_campaigns = len(items) - 1
        while i <= num_campaigns:
            j = 1
            while j < 6:
                get_result(get_user_id(items[i]), items[i], last, id_id=j)
                j += 2
            if inc_dolla == True:
                get_result(get_user_id(items[i]), items[i], last, id_id=12)
            i += 1

    # Status ID for Wrap Time Values is 1, Results by Outcome is 3, Logged In Hours is 5
    def get_init_by_agent_for_week(self, last=False, inc_dolla=True):
        def get_result(init_id, init, last, id_id):
            measurement = 0
            if last:
                st_date = LAST_MONDAY
                en_date = LAST_FRIDAY
            else:
                st_date = MONDAY
                en_date = TODAY
            self.quick_download(
                id_id, st_date, en_date, measurement, init_id=init_id, init=init
            )

        last = last
        i = 0
        items = list(OLDEST.items())
        num_users = len(items) - 1
        while i <= num_users:
            j = 1
            while j < 6:
                get_result(get_init_id(items[i][0]), items[i][0], last, id_id=j)
                j += 2
            if inc_dolla == True:
                get_result(get_init_id(items[i][0]), items[i][0], last, id_id=12)
            i += 1

    def download_results_files(self, d_start, d_end, scope=None, inc_dolla=True):
        # Get files for results updates reports from CS.
        def all_files(start_date, end_date, scope=None, inc_dolla=True):
            # Mini function that directs the chromedriver and downloads the data
            if scope == "campaigns":
                measurement = 1
            if scope == "agents":
                measurement = 0
            j = 1
            while j < 6:
                self.quick_download(j, start_date, end_date, measurement)
                j += 2
            if inc_dolla == True:
                self.quick_download(12, start_date, end_date, measurement)

        if not scope:
            all_files(d_start, d_end, "campaigns", inc_dolla)
            all_files(d_start, d_end, "agents", inc_dolla)
        if scope == "agents":
            all_files(d_start, d_end, "agents", inc_dolla)
        if scope == "campaigns":
            all_files(d_start, d_end, "campaigns", inc_dolla)

    def download_results_data(
        self, timescale=None, month=None, scope=None, inc_dolla=True
    ):
        # Download raw data to process into results report
        if timescale == "week":
            d_start = str(MONDAY)
            d_end = str(TODAY)
        elif timescale == "month":
            if month:
                d_start = str(first_of(month))
                d_end = str(end_of(month))
            else:
                d_start = str(first_of())
                d_end = str(end_of())
        else:
            d_start = str(TODAY)
            d_end = str(TODAY)
        self.download_results_files(d_start, d_end, scope, inc_dolla)

    def get_single_lr(self, init_id, folder=None, clear_folder=False):
        if not folder: folder = LR_PATH
        if clear_folder: clear_folder(folder)
        self.login()
        init = get_init_name(init_id)
        init_id = str(init_id)
        f_date = str(first_of(get_oldest_open(init)))
        t_date = str(TODAY)
        self.quick_download(report_type_id=3, f_date=f_date, t_date=t_date, measurement=5, init_id=init_id, init=init)
        self.finish_all_downloads_chrome()
        process_single_lr_file(folder=folder, file_name='download.xlsx') #### Make adaptable to clear_folder=True condition.
