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
    Contact_Export = pd.read_csv(Shira.Newest_file(r"J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA All Contacts"), encoding="latin1")
    # Assignments = pd.read_csv(Shira.Newest_file(r"J:\Admin & Plans Unit\Recovery Systems\1. Systems\Python Scripts\Morning Script\modules\Kenner_Rep\Contractor_Assignments\Test cases\Assignments"), encoding="latin1")
    Contractors_unassigned = Contact_Export[["Email","First Name","Last Name","Group(s)","Applicant?", "Item Link"]]   


    Contractors_unassigned = Contractors_unassigned[Contractors_unassigned["Group(s)"] == "Contractor"]

    Contractors_unassigned = Contractors_unassigned[Contractors_unassigned["Applicant?"] == "N"]

    Contractors_unassigned = Contractors_unassigned[["Email"]]
  


    Assignments = Assignments[Assignments["Contact Group(s)"] == "Contractor"]

    Assignments = Assignments[["Grant", "County", "Applicant", "Position(s)", "Contact Email"]]

    Contractors_unassigned = Contractors_unassigned[~Contractors_unassigned["Email"].isin(Assignments["Contact Email"])]




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
        "EY": "Nadya.Semenova@ey.com",
        "Deloitte": "chbreed@deloitte.com",
        "Horne": "sam.hurst@horne.com",
        "DCMC": "mrobinson@dcmcpartners.com",
        "Tidal_Basin": "asupriana@tidalbasin.rphc.com",
        "THF": "Bbechtel@thf-cpa.com",
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

            # match = current_assignments[["Grant", "County", "Applicant"]].equals(default_assignments[["Grant", "County", "Applicant"]])
            # finding missing rows
            missing_rows = pd.concat([default_assignments, current_assignments]).drop_duplicates(subset=['Grant', 'County', 'Applicant', "Position(s)"], keep=False)
            # Count missing rows
            missing_count = len(missing_rows)

            if missing_count == 0:
                # print (f"{value} has correct assignments")
                continue
            else:
                change_list.append(value)
                # print(f"{value} had {missing_count} wrong assignments")


    def execute_funcations():
        for email in company_emails.keys():
            employee_df = contractor_employees(email)
            assignment_df = account_assignments(email)
            change_log(employee_df,assignment_df)

    def map_assignment_emails(change_list, company_emails, assignment_accounts):
        # Ensure change_list is a DataFrame
        if isinstance(change_list, list):
            change_list = pd.DataFrame(change_list, columns=["Email"])
        
        # Create an empty column for default assignment emails
        change_list["Default Assignment Email"] = "Not Found"

        # Loop through each email and find the matching assignment email
        for index, row in change_list.iterrows():
            for company, domain in company_emails.items():
                if domain in row["Email"]:  # Check if the domain exists in the email
                    change_list.at[index, "Default Assignment Email"] = assignment_accounts.get(company, "Not Found")
                    break  # Stop searching once a match is found
        return change_list

    execute_funcations()

    change_list = pd.DataFrame(change_list, columns=["Email"])

    change_list = pd.concat([change_list, Contractors_unassigned], ignore_index=True)

    change_list.to_csv(r"J:\Admin & Plans Unit\Recovery Systems\1. Systems\Python Scripts\Morning Script\modules\Kenner_Rep\Contractor_Assignments\Assignment Changes\Pending_Changes.csv", index=False)


    change_list = map_assignment_emails(change_list, company_emails, assignment_accounts)

    return change_list


if __name__ == "__main__":
    main()