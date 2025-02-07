def main(pending_changes):
    import time as t
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys
    from selenium.common.exceptions import TimeoutException
    from selenium.webdriver.support.ui import Select
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
    import Functions
    from Functions import copy_assignments, delete_assignments, FLPA_sign_in

    options=webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # options.add_argument("--headless")
    # options.add_argument("--disable-software-rasterizer")
    driver_service=Service(r'J:\Admin & Plans Unit\Recovery Systems\1. Systems\ChromeDriver\chromedriver.exe')
    driver=webdriver.Chrome(service=driver_service, options=options)
    wait=WebDriverWait(driver, 120)

    unassigned_users = pd.DataFrame(columns=["Name"])

    Contact_Export = pd.read_csv(Shira.Newest_file(r"J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA All Contacts"), encoding="latin1")

    # Contact_Export = pd.read_csv(Shira.Newest_file(r"J:\Admin & Plans Unit\Recovery Systems\1. Systems\Python Scripts\Morning Script\modules\Kenner_Rep\Contractor_Assignments\Test cases\Contacts"), encoding="latin1")

    Contact_Export = Contact_Export.applymap(lambda x: x.strip().lower() if isinstance(x, str) else x)
    pending_changes = pending_changes.applymap(lambda x: x.strip().lower() if isinstance(x, str) else x)


    Contact_Export = Contact_Export[["Email","First Name","Last Name", "Item Link"]]

    Contact_Export["Name"] = Contact_Export["First Name"] + " " + Contact_Export["Last Name"]

    Contact_Export = Contact_Export[["Email","Name", "Item Link"]]

    pending_changes = pending_changes.merge(Contact_Export[['Email', 'Item Link']], left_on="Default Assignment Email",right_on="Email", how="left")


    pending_changes.drop(columns=["Email_y"], inplace=True)
    pending_changes.rename(columns={"Item Link": "Assignments Link"}, inplace=True)
    pending_changes.rename(columns={"Email_x": "Email"}, inplace=True)

    pending_changes = pending_changes.merge(Contact_Export, on="Email", how="left")
    pending_changes["Item Link"] = pending_changes["Item Link"] + "?t=form--assignments"
    pending_changes["Assignments Link"] = pending_changes["Assignments Link"] + "?t=form--assignments"

    pending_changes = pending_changes.dropna(subset=["Assignments Link"])

    print(f"There are ({len(pending_changes)}) Contractors to update")

    pending_changes.to_csv(r"J:\Admin & Plans Unit\Recovery Systems\1. Systems\Python Scripts\Morning Script\modules\Kenner_Rep\Contractor_Assignments\Assignment Changes\Pending_Changes.csv", index=False)

    Functions.FLPA_sign_in(driver)



    for index, row in pending_changes.iterrows():
        Link = row["Item Link"]
        Contact_Name = row["Name"]
        Email = row["Email"]
        assignments = row["Assignments Link"]
        try:
            driver.refresh()
            t.sleep(3)
            try:
                Functions.delete_assignments(driver, Link,Contact_Name, unassigned_users)
                Deleted = True
            except:
                Deleted = False
            if Deleted == True:
                try:
                    t.sleep(2)
                    # print("test")
                    Functions.copy_assignments(driver,assignments,Contact_Name,Email)
                    unassigned_users = unassigned_users[unassigned_users["Name"] != f"{Contact_Name}"]
                    t.sleep(5)
                except:
                    t.sleep(2)
                    # print("test")
                    Functions.copy_assignments(driver,assignments,Contact_Name,Email)
                    unassigned_users = unassigned_users[unassigned_users["Name"] != f"{Contact_Name}"]
                    t.sleep(5)
            else:
                print (f"failed to delete {Contact_Name}'s assignments")
        except:
            driver.refresh()
            print(f"failed to edit {Contact_Name}. Ensure they still have assignments")
    
    if len(unassigned_users) > 0:
        Shira.webex_bot(message="Contractors have been left without assignments please see the list")
        Shira.webex_bot(message=f"{unassigned_users}")
        print(f"{unassigned_users}")
    else:
        print("No user's with empty assignment list")

    



if __name__ == "__main__":
    main()