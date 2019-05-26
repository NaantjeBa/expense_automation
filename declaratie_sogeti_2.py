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
        Lets user input month, checks if input is a type int and returns it.
        :return: The value of the input month
        :rtype: int
        """

        month_nr = ''
        while type(month_nr) != int:
            try:
                month_nr = input('Fill in expense month: ')
                month_nr = int(month_nr)
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
        last_day = month_range[1]
        from_date = datetime.datetime(year, month_nr, 1)
        until_date = datetime.datetime(year, month_nr, last_day)

        return from_date, until_date

    def get_excelfile_ns(from_date, until_date):
        """
        Takes in first and last date, opens browser and downloads excel file from NS website.
        :param from_date: The first date of the period
        :type from_date: datetime
        :param until_date: The last date of the period
        :type until_date: datetime
        :return: None
        """

        # username = 'jesse.niens@sogeti.com'
        # password = 'D1hbwvN!'

        from_day_str = str(from_date.day).zfill(2)
        from_month_str = str(from_date.month).zfill(2)
        from_year_str = str(from_date.year).zfill(4)

        until_day_str = str(until_date.day).zfill(2)
        until_month_str = str(until_date.month).zfill(2)
        until_year_str = str(until_date.year).zfill(4)

        # Open browser and go to ns.nl login page
        os.chdir('C:\\Users\\jniens\\Downloads')
        browser = webdriver.Chrome()
        browser.get('https://www.ns.nl/mijnnszakelijk/login?0')

        # Pause the program to go to let user log in and navigate to fill in page
        os.system('pause')

        #klik gemaakte reizen
        gemaakte_reizen = browser.find_element_by_xpath('//*[@id="menuitem.label.hybristravelhistory"]')
        gemaakte_reizen.click()

        from_day = browser.find_element_by_xpath('//*[@id="dayField"]')
        from_day.clear()
        from_day.send_keys(from_day_str)
        from_month = browser.find_element_by_xpath('//*[@id="monthField"]')
        from_month.clear()
        from_month.send_keys(from_month_str)
        from_year = browser.find_element_by_xpath('//*[@id="yearField"]')
        from_year.clear()
        from_year.send_keys(from_year_str)
        actionchains = ActionChains(browser)
        actionchains.send_keys(Keys.TAB)
        actionchains.send_keys(Keys.TAB)
        actionchains.send_keys(until_day_str)
        actionchains.send_keys(until_month_str)
        actionchains.send_keys(until_year_str)
        actionchains.perform()

        # Click the Zoeken button
        button_zoeken = browser.find_element_by_xpath('/ html / body / main / div / div / div / div / div / div[2] / div[2] / div[1] / form / p / a[1] / span')
        button_zoeken.click()
        # download the excel file
        button_download = browser.find_element_by_css_selector('#ns-app > div.col-3b > div.title.box > ul > li > a')
        button_download.click()

        # "Pause until download is finished"
        # os.system('pause')

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
            print("Input amount and calculated amount from expense report match!\n"
                  "Continuing...")
        else: # If amounts don't match inform user and give option to abort or continue anyway
            print(f"Input amount and calculated amount don't match:\n"
                  f"\n"
                  f"Input amount: {input_amount}\n"
                  f"Calculated amount: {calc_amount_round}\n")

            while True:
                value = input('Do you want to continue anyway [y to continue / n to quit]? ')
                if value.lower() == 'y':
                    print('Continuing...')
                    break
                elif value.lower() == 'n':
                    print('Abort process...')
                    exit()
                else:
                    print('Please input y or n')


        return df

    def open_browser_sogeti(from_date, input_amount):
        """
        Reads in first date and amount, opens browser and fills in the expenses basics (including amount).

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

        # Open browser and go to einstein.sogeti.nl
        os.chdir('C:\\Users\\jniens\\Downloads')
        browser = webdriver.Chrome()
        browser.get('https://einstein.sogeti.nl/')

        # Pause the program to go to let user log in and navigate to fill in page
        os.system('pause')

        #mijnSogetibutton
        browser.find_element_by_xpath('//*[@id="block-menu-block-2"]/div/div/ul/li[2]/a').click()

        #mijnDeclaratie
        browser.find_element_by_xpath('//*[@id="block-menu-block-5"]/div/div/ul/li[5]/ul/li[1]/a').click()

        # Find the iFrame on fill in page
        frame = browser.find_element_by_xpath('//*[@id="node-34"]/div/div/div/div/iframe')
        browser.switch_to.frame(frame)

        # dropdown Reiskosten YP
        dropdown = Select(browser.find_element_by_xpath('/html/body/form/table/tbody/tr[4]/td/select'))
        dropdown.select_by_visible_text('Reiskosten YP')

        # Press "Verder"
        browser.find_element_by_xpath('//*[@id="verderButton"]').click()

        #Mijn referentie
        txt_box_ref = browser.find_element_by_xpath('/html/body/form/table/tbody/tr[8]/td[2]/input')
        txt_box_ref.send_keys(mijn_referentie)

        #Bedrag
        txt_box_amount = browser.find_element_by_xpath('/html/body/form/table/tbody/tr[10]/td[2]/input')
        txt_box_amount.send_keys(input_amount)

        # Click "vervolg declaratie
        browser.find_element_by_xpath('//*[@id="bvzm"]').click()

        # # Change iframe
        # frame = browser.find_element_by_xpath('//*[@id="node-34"]/div/div/div/div/iframe')
        # browser.switch_to.frame(frame)

        # Fill in month nr
        txt_box_monthnr = browser.find_element_by_xpath('//*[@id="decHeadings[0].decHeadingsValue"]')
        txt_box_monthnr.send_keys(month_nr)
        txt_box_monthnr.send_keys(Keys.TAB)
        actionchains = ActionChains(browser)
        actionchains.send_keys(year)
        actionchains.send_keys(Keys.TAB)
        actionchains.send_keys(2)
        actionchains.perform()

        return browser

    def loop_through_df(df, browser):
        """
        Fills in the expenses row by row on the webpage.

        :param df: The dataframe containing the expenses
        :type df: dataframe
        :param browser: The driver object to be used to fill in the expenses row by row
        :type browser: WebDriver
        :return: None
        """
        # nr_columns = len(df.columns)
        rows = len(df)

        "Modify for variable i based on variable x for inconsistencies in xpath names"
        for x, row in df.iterrows():
            if x == 0:
                i = 0
                input_factuurdatum = browser.find_element_by_xpath('//*[@id="' + str(i) + '_2"]')
            elif x == 1:
                i = 2
                input_factuurdatum = browser.find_element_by_name(str(i) + '_2')
            else:
                i = x + 1
                input_factuurdatum = browser.find_element_by_xpath('//*[@id="' + str(i) + '_2"]')

            "Extract datum, prev_datum and ov_bedrag"
            datum = row.Datum
            prev_datum = df.iloc[x - 1, 1]
            #datum_str = datum.strftime('%d-%m-%Y')
            ov_bedrag = str(row['Prijs (incl. btw)'])

            "Determine 'ritnummer' based on datum and prev_datum"
            if x == 0 or prev_datum != datum:
                rit_nummer = 1
            else:
                rit_nummer += 1

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

            # Fill in the values in online form
            input_factuurdatum.send_keys(datum)
            actionchains = ActionChains(browser)
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

            # Press voeg_lege_regel_toe or opslaan_controle at last row
            if x < rows - 1:
                voeg_lege_regel_toe = browser.find_element_by_css_selector(
                    'body > form > table:nth-child(3) > tbody > tr:nth-child(2) > td > input.button')
                voeg_lege_regel_toe.click()
            else:
                opslaan_controle = browser.find_element_by_css_selector(
                    'body > form > table:nth-child(3) > tbody > tr:nth-child(2) > td > input:nth-child(13)')
                opslaan_controle.click()

        # Pause the program, let user decide if he wants to progress
        os.system('pause')

        # Unselect all rows
        for i in range(rows):
            vinkje = browser.find_element_by_id('regelcheck' + str(i + 1))
            vinkje.click()

        # Press OpslaanControle
        opslaan_controle = browser.find_element_by_css_selector(
            'body > form > table:nth-child(3) > tbody > tr:nth-child(2) > td > input:nth-child(13)')
        opslaan_controle.click()

    year = input_user_year()
    month_nr = input_user_month()
    amount = input_user_amount()
    from_date, until_date = define_period(year, month_nr)
    get_excelfile_ns(from_date, until_date)
    df = read_in_df()
    # df = pd.read_excel('C:\\Users\\jniens\\Downloads\\reistransacties-3528010488672904 (16).xls')
    df = filter_out_zero(df)
    check_amount(df, amount)
    browser = open_browser_sogeti(from_date, amount)
    loop_through_df(df, browser)


if __name__ == '__main__':
    main()
