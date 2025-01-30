def main():
    import time
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import TimeoutException
    from zipfile import ZipFile
    import shutil
    import keyring
    from datetime import date, datetime, timedelta
    import win32com.client
    import pandas as pd
    import calendar
    import os
    import re

    import Shira
    from Shira import Newest_file

    change_list = []

    Assignments = pd.read_csv(Shira.Newest_file(r"J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Assignments"), encoding="latin1")

    Assignments = Assignments[Assignments["Contact Group(s)"] == "Contractor"]

    Assignments = Assignments[["Grant", "County", "Applicant", "Position(s)", "Contact Email"]]

    print(Assignments)

    # define what email addresses to look for
    company_emails = {
        "Atkins": "@atkinsrealis.com",
        "Hagerty": "@hagertyconsulting.com",
        "KPMG": "@kpmg.com",
        "RSM": "@rsmus.com",
        "IEM": "@iem.com",
        "EY": "@ey.com",
        "Deloitte": "@deloitte.com",
        "Horne": "@horne.com",
        "DCMC": "@dcmcpartners.com",
        "Tidal_Basin": "@tidalbasin.rphc.com",
        "THF": "@thf-cpa.com",
        "CRI": "@cricpa.com"
    }
    # define where assignments should come from
    assignment_accounts = {
        "Atkins": "jamelyn.trucks@atkinsrealis.com",
        "Hagerty": "danielle.finella@hagertyconsulting.com",
        "KPMG": "samanthasicard@kpmg.com",
        "RSM": "regina.oliver@rsmus.com",
        "IEM": "shaun.mcgrath@iem.com",
        "EY": "nadya.semenova@ey.com",
        "Deloitte": "chbreed@deloitte.com",
        "Horne": "sam.hurst@horne.com",
        "DCMC": "mrobinson@dcmcpartners.com",
        "Tidal_Basin": "rwright@tidalbasin.rphc.com",
        "THF": "bbechtel@thf-cpa.com",
        "CRI": "lluong@cricpa.com"
    }

    # get the list of employees
    def contractor_employees(company_name):
        if company_name in company_emails:
            domain = company_emails[company_name]
            employees = Assignments[Assignments['Contact Email'].str.contains(domain, na=False)]
            return employees
        else:
            print(f"Company '{company_name}' not found in dictionary.")
        return None
    
    # get the default assignment for the contractor
    def account_assignments(company_name):
        if company_name in assignment_accounts:
            default_assignments = Assignments[Assignments["Contact Email"] == assignment_accounts[company_name]]
            return default_assignments
        else:
            print(f"Company '{company_name}' not found in dictionary.")
        return None

    # make a log of all the users to be changed
    def change_log(employees,default_assignments):
        induviduals = employees['Contact Email'].unique()

        for value in induviduals:
            current_assignments = Assignments[Assignments["Contact Email"] == value]

            match = current_assignments[["Grant", "County", "Applicant"]].equals(default_assignments[["Grant", "County", "Applicant"]])

            if match:
                print (f"{value} has correct assignments")
            else:
                change_list.append(value)


    def execute_funcations():
        for email in company_emails.keys():
            employee_df = contractor_employees(email)
            assignment_df = account_assignments(email)
            change_log(employee_df,assignment_df)

    execute_funcations()

    return change_list


if __name__ == "__main__":
    main()