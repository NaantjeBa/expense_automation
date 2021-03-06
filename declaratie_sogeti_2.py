#! python3
"""This program loops through public transport expenses and fills in
the online form on Einstein.sogeti.nl"""
import calendar
import datetime
import re
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import glob
import pandas as pd
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException


def main():
    """This is the main function"""

    def input_user_year():
        """
        Lets user input year, checks if input is a type int and returns it.

        :return: The value of the input year
        :rtype: int
        """

        year = ''
        while type(year) != int:
            try:
                year = input('Fill in expense year: ')
                year = int(year)
            except ValueError:
                print("Please fill in an integer")

        return year

    def input_user_month():
        """
        Lets user input month, checks if input is a type int and between 0 and 13. Then returns it.

        :return: The value of the input month
        :rtype: int
        """

        month_nr = ''
        within_range = False
        while type(month_nr) != int or not within_range:
            try:
                month_nr = input('Fill in expense month (between 0 and 13): ')
                month_nr = int(month_nr)
                within_range = 0 < month_nr < 13
            except ValueError:
                print("Please fill in an integer")

        return month_nr

    def input_user_amount():
        """
        Lets user input amount, checks if input is a type float and returns it.

        :return: The value of the input amount
        :rtype: float
        """

        amount = ''
        while type(amount) != float:
            try:
                amount  = input('Fill in expense amount: ')
                amount = float(amount)
            except ValueError:
                print("Please fill in a number")

        return amount

    def define_period(year, month_nr):
        """
        Takes in year and month and returns first and last date.

        :param year: The year value
        :type year: int
        :param month_nr: The month value
        :type month_nr: int
        :return: tuple (from_date, until_date)
            WHERE
            datetime from_date is the first date of the period
            datetime until_date is the last date of the period
        """

        month_range = calendar.monthrange(year, month_nr)
        first_day = 1
        last_day = month_range[1]
        from_date = datetime.datetime(year, month_nr, first_day)
        until_date = datetime.datetime(year, month_nr, last_day)

        return from_date, until_date

    def string_period(from_date, until_date):
        """
        Turns datetime objects into string objects to fill in on NS webpage.

        :param from_date: from_date: The first date of the period
        :type from_date: datetime
        :param until_date: The last date of the period
        :type until_date: datetime
        :return: A dictionary containing the string-date values to fill in on NS webpage.
        :rtype: dict
        """
        date_dict_str = dict()
        date_dict_str['from_day'] = str(from_date.day).zfill(2)
        date_dict_str['from_month'] = str(from_date.month).zfill(2)
        date_dict_str['from_year'] = str(from_date.year).zfill(4)

        date_dict_str['until_day'] = str(until_date.day).zfill(2)
        date_dict_str['until_month'] = str(until_date.month).zfill(2)
        date_dict_str['until_year'] = str(until_date.year).zfill(4)

        return date_dict_str

    def login_ns_webpage():
        """
        Opens browser and aks user to login to NS webpage.

        :return: the webbrowser where user is logged in
        :rtype: WebDriver
        """

        # Open browser and go to ns.nl login page
        os.chdir('C:\\Users\\jniens\\Downloads')
        browser_ns = webdriver.Chrome()
        browser_ns.get('https://www.ns.nl/mijnnszakelijk/login?0')

        "Print instruction message for user"
        print("---------------------------------------------------------------\n"
              "Please log in to NS webpage using your credentials")

        "Pause the program to go to let user log in and navigate to fill in page"
        os.system('pause')

        return browser_ns

    def check_ns_element(browser_ns):
        """
        Checks if element "gemaakte reizen" is found and clicks it.

        :param browser_ns: The browser where the user is supposed to be logged in.
        :return: None
        """

        "Try to find 'gemaakte reizen' and click it"
        gemaakte_reizen = None
        while not gemaakte_reizen:
            try:
                gemaakte_reizen = browser_ns.find_element_by_xpath('//*[@id="menuitem.label.hybristravelhistory"]')
                gemaakte_reizen.click()
            except NoSuchElementException:
                print("---------------------------------------------------------------\n"
                      "Dit not find 'Gemaakte reizen'\n"
                      "Please make sure you are logged in to the NS webpage and see the 'Gemaakte Reizen' element\n"
                      "Then press Enter")
                os.system('pause')

    def download_excel_file(date_dict_str, browser_ns):
        """
        Takes in browser and dictionary containing date strings and downloads excel file.

        :param date_dict_str: A dictionary containing the dates to fill in on ns webpage in string format
        :type date_dict_str: dict
        :param browser_ns: The browser object where the user is logge in to NS webpage
        :type browser_ns: WebDriver
        :return:
        """

        "Get string values to fill in on NS webpage"
        from_day = browser_ns.find_element_by_xpath('//*[@id="dayField"]')
        from_day.clear()
        from_day.send_keys(date_dict_str['from_day'])
        from_month = browser_ns.find_element_by_xpath('//*[@id="monthField"]')
        from_month.clear()
        from_month.send_keys(date_dict_str['from_month'])
        from_year = browser_ns.find_element_by_xpath('//*[@id="yearField"]')
        from_year.clear()
        from_year.send_keys(date_dict_str['from_year'])
        actionchains = ActionChains(browser_ns)
        actionchains.send_keys(Keys.TAB)
        actionchains.send_keys(Keys.TAB)
        actionchains.send_keys(date_dict_str['until_day'])
        actionchains.send_keys(date_dict_str['until_month'])
        actionchains.send_keys(date_dict_str['until_year'])
        actionchains.perform()

        "Click the Zoeken button"
        button_zoeken = browser_ns.find_element_by_xpath('/ html / body / main / div / div / div / div / div / div[2] / div[2] / div[1] / form / p / a[1] / span')
        button_zoeken.click()

        "Download the excel file"
        button_download = browser_ns.find_element_by_css_selector('#ns-app > div.col-3b > div.title.box > ul > li > a')
        button_download.click()

        "Inform user to wait until download is finished"
        print("---------------------------------------------------------------\n"
              "Please wait until download is finished")

        "Pause until download is finished"
        os.system('pause')

        "Close NS browser"
        browser_ns.close()

    def read_in_df():
        """
        Reads in downloaded excel file and returns it as a dataframe.

        :return: The dataframe containing the expenses
        :rtype: dataframe
        """
        list_of_files = glob.glob('C:\\Users\\jniens\\Downloads\\*')
        latest_file = max(list_of_files, key=os.path.getctime)
        df = pd.read_excel(latest_file)
        # df = pd.read_excel('C:\\Users\\jniens\\Downloads\\reistransacties-3528010488672904 (11).xls')

        return df

    def filter_out_zero(df):
        """
        Cleans the dataframe and returns it.

        :param df: The dataframe containing the expenses
        :type df: dataframe
        :return: The cleaned dataframe containing the expenses
        :rtype: dataframe
        """
        df.drop(df.tail(1).index, inplace=True)
        df = df[df["Prijs (incl. btw)"] != 0]
        df = df.sort_values('Datum')
        df.reset_index(inplace=True)
        del df['index']

        return df

    def check_amount(df, input_amount):
        """
        Check whether inputted amount and calculated amount from Excel match and gives user option to continue if
        values don't.

        :param df: The dataframe containing the expenses
        :type df: dataframe
        :param input_amount: The input amount from user
        :type input_amount: float
        :return: None
        """

        # Calculate total amount
        calc_amount = df['Prijs (incl. btw)'].sum()

        # Round total amount to 2 decimals
        calc_amount_round = round(calc_amount, 2)

        # Check if input amount and calculated amount match
        if calc_amount_round == input_amount:  # If both amounts match give user confirmation and continue
            print("---------------------------------------------------------------\n"
                  "Input amount and calculated amount from expense report match!\n"
                  "Continuing...")
        else: # If amounts don't match inform user and give option to abort or continue anyway
            print(f"---------------------------------------------------------------\n"
                  "Input amount and calculated amount don't match:"
                  f"\n"
                  f"Input amount: {input_amount}\n"
                  f"Calculated amount: {calc_amount_round}")

            while True:
                value = input('Do you want to continue anyway [y to continue / n to quit]? ')
                if value.lower() == 'y':
                    print('Continuing...')
                    break
                elif value.lower() == 'n':
                    print("---------------------------------------------------------------\n"
                          'Abort process...')
                    exit()
                else:
                    print("---------------------------------------------------------------\n"
                          'Please input y or n')


        return df

    def login_sogeti_webpage():
        """
       Opens browser and aks user to login to Sogeti webpage.

        :return: the webbrowser where user is logged in
        :rtype: WebDriver
        """

        # Open browser and go to ns.nl login page
        os.chdir('C:\\Users\\jniens\\Downloads')
        browser_sogeti = webdriver.Chrome()
        browser_sogeti.get('https://einstein.sogeti.nl/')

        "Print instruction message for user"
        print("---------------------------------------------------------------\n"
              "Please log in to Sogeti webpage using your credentials")

        "Pause the program to go to let user log in and navigate to fill in page"
        os.system('pause')

        return browser_sogeti

    def check_sogeti_element(browser_sogeti):
        """
        Checks if element "Mijn Sogeti" is found and clicks it.

        :param browser_sogeti: The browser where the user is supposed to be logged in.
        :type browser_sogeti: WebDriver
        :return: None
        """

        "Try to find 'gemaakte reizen' and click it"
        mijn_sogeti = None
        while not mijn_sogeti:
            try:
                mijn_sogeti = browser_sogeti.find_element_by_xpath('//*[@id="block-menu-block-2"]/div/div/ul/li[2]/a')
                mijn_sogeti.click()
            except NoSuchElementException:
                print("---------------------------------------------------------------\n"
                      "Dit not find 'Gemaakte reizen'\n"
                      "Please make sure you are logged in to the Sogeti webpage and see the 'Mijn Sogeti' element\n"
                      "Then press Enter")
                os.system('pause')

    def fill_in_basics_sogeti(browser_sogeti, from_date, input_amount):
        """
        Reads in first date and amount, opens browser and fills in the expenses basics (including amount).

        :param browser_sogeti: The webbrowser where user is logged in to Sogeti webpage.
        :type browser_sogeti: WebDriver
        :param from_date: first date of the period
        :type from_date: datetime
        :param input_amount: The inputted amount of the user
        :type input_amount: float
        :return: The driver object to be used to fill in the expenses row by row
        :rtype: WebDriver
        """

        month_nr = from_date.month
        year = from_date.year
        month_and_year = from_date.strftime('%B %Y')
        mijn_referentie = f'Expenses for {month_and_year}'
        input_amount = str(input_amount)

        # mijnDeclaratie
        browser_sogeti.find_element_by_xpath('//*[@id="block-menu-block-5"]/div/div/ul/li[5]/ul/li[1]/a').click()

        # Find the iFrame on fill in page
        frame = browser_sogeti.find_element_by_xpath('//*[@id="node-34"]/div/div/div/div/iframe')
        browser_sogeti.switch_to.frame(frame)

        # dropdown Reiskosten YP
        dropdown = Select(browser_sogeti.find_element_by_xpath('/html/body/form/table/tbody/tr[4]/td/select'))
        dropdown.select_by_visible_text('Reiskosten YP')

        # Press "Verder"
        browser_sogeti.find_element_by_xpath('//*[@id="verderButton"]').click()

        #Mijn referentie
        txt_box_ref = browser_sogeti.find_element_by_xpath('/html/body/form/table/tbody/tr[8]/td[2]/input')
        txt_box_ref.send_keys(mijn_referentie)

        #Bedrag
        txt_box_amount = browser_sogeti.find_element_by_xpath('/html/body/form/table/tbody/tr[10]/td[2]/input')
        txt_box_amount.send_keys(input_amount)

        # Click "vervolg declaratie
        browser_sogeti.find_element_by_xpath('//*[@id="bvzm"]').click()

        # # Change iframe
        # frame = browser.find_element_by_xpath('//*[@id="node-34"]/div/div/div/div/iframe')
        # browser.switch_to.frame(frame)

        # Fill in month nr
        txt_box_monthnr = browser_sogeti.find_element_by_xpath('//*[@id="decHeadings[0].decHeadingsValue"]')
        txt_box_monthnr.send_keys(month_nr)
        txt_box_monthnr.send_keys(Keys.TAB)
        actionchains = ActionChains(browser_sogeti)
        actionchains.send_keys(year)
        actionchains.send_keys(Keys.TAB)
        actionchains.send_keys(2)
        actionchains.perform()

        return browser_sogeti

    def loop_through_df(df, browser_sogeti):
        """
        Fills in the expenses row by row on the webpage.

        :param df: The dataframe containing the expenses
        :type df: dataframe
        :param browser_sogeti: The driver object to be used to fill in the expenses row by row
        :type browser_sogeti: WebDriver
        :return: None
        """

        def find_element(index):
            """
            Finds the element name of the input textbox based on the index of the dataframe.

            :param index: the index of the row in the dataframe
            :type index: int
            :return: the name of the element to fill in the information
            :rtype: str
            """

            if index != 0:
                index += 1
            element_name = str(index) + "_2"

            return element_name

        def return_date():
            """
            Returns the date of the row.

            :return: the date of the expense
            :rtype: str
            """
            datum = row.Datum

            return datum

        def return_ovbedrag():
            """
            Returns the amount of the expense row.

            :return: the expense amount
            :rtype: str
            """
            ov_bedrag = row['Prijs (incl. btw)']
            ov_bedrag_str = str(ov_bedrag)

            return ov_bedrag_str

        def return_ritnummer(rit_nummer):
            """
            Returns the ritnummer to fill in on online form.

            :param rit_nummer: The rit_nummer of the previous row.
            :type rit_nummer: int
            :return: The rit_nummer for the current row
            :rtype: int
            """
            prev_datum = df.iloc[index - 1, 1]
            datum = return_date()

            "Determine 'ritnummer' based on datum and prev_datum"
            if index == 0 or prev_datum != datum:
                rit_nummer = 1
            else:
                rit_nummer += 1

            return rit_nummer

        def return_van_naar():
            """
            Returns the from and to locations of the expense row.

            :return:
            """
            "Determine van_halte en naar_halte"
            omschrijving = row.Omschrijving
            find_cor = omschrijving.find('Correctietarief:')  # Checks if expense-record is Correctietarief

            find_uit = omschrijving.find('-uit:')

            "Find at what index to slice string"
            if find_uit > -1:
                start_string = find_uit + 6
            else:
                start_string = 0

            sliced_str = omschrijving[start_string:]  # Slice string
            find_sep = sliced_str.find('-')  # Returns -1 if "-" is not found

            "This part extracts start and stop for each record" \
            "If Exception is raised it gives default values"
            van_halte = "Vanaf halte/station"
            naar_halte = "Naar halte/station"
            try:
                if find_cor != -1:
                    van_halte = "Correctietarief"
                    naar_halte = "Correctietarief"

                elif find_sep > 0:
                    van_halte = sliced_str[:find_sep - 1]
                    naar_halte = sliced_str[find_sep + 2:]

                else:
                    halte_regex = re.compile(r'(halte(\s[A-Z]\w+(\.)?)*)')
                    van_halte = halte_regex.findall(sliced_str)[0][0]
                    naar_halte = halte_regex.findall(sliced_str)[1][0]

            except:
                pass

            return van_halte, naar_halte

        def generate_element():
            """
            Generates browser element based on string name.

            :return: The element to fill in the first value (factuurdatum) of the expense row.
            :rtype: WebElement
            """

            element = browser_sogeti.find_element_by_name(element_name)

            return element

        def fill_in_values():
            """
            Fills in on values in in online form.

            :return:
            """
            # Fill in the values in online form
            element.send_keys(datum)
            actionchains = ActionChains(browser_sogeti)
            actionchains.send_keys(Keys.TAB)
            actionchains.send_keys(rit_nummer)
            actionchains.send_keys(Keys.TAB)
            actionchains.send_keys(ov_bedrag)
            actionchains.send_keys(Keys.TAB)
            actionchains.send_keys(Keys.TAB)
            actionchains.send_keys(van_halte)
            actionchains.send_keys(Keys.TAB)
            actionchains.send_keys(naar_halte)
            actionchains.perform()

        def press_button():
            """
            Decides whether to add new row or not based on which iteration it is on.

            :return:
            """
            # Press voeg_lege_regel_toe or opslaan_controle at last row
            voeg_lege_regel_toe = browser_sogeti.find_element_by_css_selector(
                'body > form > table:nth-child(3) > tbody > tr:nth-child(2) > td > input.button')
            try:
                opslaan_controle = browser_sogeti.find_element_by_css_selector(
                'body > form > table:nth-child(3) > tbody > tr:nth-child(2) > td > input:nth-child(13)')
            except NoSuchElementException:
                print('Niet gevonden')

            if index < rows - 1:
                voeg_lege_regel_toe.click()
            else:
                opslaan_controle.click()

        def unselect_all():
            """
            Unselects all expenses on online form.

            :return:
            """
            opslaan_controle = browser_sogeti.find_element_by_css_selector(
                'body > form > table:nth-child(3) > tbody > tr:nth-child(2) > td > input:nth-child(13)')

            # Unselect all rows
            for i in range(rows):
                vinkje = browser_sogeti.find_element_by_id('regelcheck' + str(i + 1))
                vinkje.click()

            # Press OpslaanControle
            opslaan_controle.click()

        rows = len(df)
        rit_nummer = 0
        "Modify for variable i based on variable index for inconsistencies in xpath names"
        for index, row in df.iterrows():

            element_name = find_element(index)
            datum = return_date()
            ov_bedrag = return_ovbedrag()
            rit_nummer = return_ritnummer(rit_nummer)
            van_halte, naar_halte = return_van_naar()
            element = generate_element()
            fill_in_values()
            press_button()

        # Pause the program, let user decide if he wants to progress
        os.system('pause')
        unselect_all()

    year = input_user_year()
    month_nr = input_user_month()
    amount = input_user_amount()
    from_date, until_date = define_period(year, month_nr)
    date_dict_str = string_period(from_date, until_date)
    # browser_ns = login_ns_webpage()
    # check_ns_element(browser_ns)
    # download_excel_file(date_dict_str, browser_ns)
    df_raw = read_in_df()
    # df_raw = pd.read_excel('C:\\Users\\jniens\\Downloads\\reistransacties-3528010488672904 (16).xls')
    df_filtered = filter_out_zero(df_raw)
    check_amount(df_filtered, amount)
    browser_sogeti = login_sogeti_webpage()
    check_sogeti_element(browser_sogeti)
    browser_sogeti = fill_in_basics_sogeti(browser_sogeti, from_date, amount)
    loop_through_df(df_filtered, browser_sogeti)


if __name__ == '__main__':
    main()
