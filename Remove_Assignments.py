def main(pending_changes):
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

    Contact_Export = pd.read_csv(Shira.Newest_file(r"J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA All Contacts"), encoding="latin1")

    Contact_Export = Contact_Export[["Email","First Name","Last Name", "Item Link"]]

    Contact_Export["Name"] = Contact_Export["First Name"] + " " + Contact_Export["Last Name"]

    Contact_Export = Contact_Export[["Email","Name", "Item Link"]]

    pending_changes = pending_changes.merge(Contact_Export, on="Email", how="left")

    pending_changes.to_csv(r"J:\Admin & Plans Unit\Recovery Systems\1. Systems\Python Scripts\Morning Script\modules\Kenner_Rep\Contractor_Assignments\Assignment Changes\Pending_Changes.csv", index=False)


if __name__ == "__main__":
    main()