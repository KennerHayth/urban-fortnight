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

def FLPA_sign_in(driver):
    FLPA_GP_username=keyring.get_password("FLPA_GP", "username")
    FLPA_password=keyring.get_password("FLPA", "FLPA_password")
    driver.get(r"https://floridapa.org/")
    t.sleep(4)
    username_field=driver.find_element(By.NAME,"Username")
    password_field=driver.find_element(By.NAME,"Password")
    username_field.clear()
    password_field.clear()
    username_field.send_keys(FLPA_GP_username)
    password_field.send_keys(FLPA_password)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "Submit"))
    ).click()
    t.sleep(8)

def select_all_assignments(driver):
    dropdown = Select(driver.find_element(By.XPATH, "/html/body/div[6]/div[2]/div/div[2]/form/fieldset/div[6]/select"))
    
    
    # Loop through all available options and select each one
    for option in dropdown.options:
        dropdown.select_by_visible_text(option.text)


def select_first_matching_email(driver, target_email):
    try:
        # Wait for elements to be present
        elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "/html/body/div[6]/div[2]/div/div[2]/form/fieldset/div[2]/div/div/div/ul/li"))
        )

        # Filter matching elements
        matching_elements = [element for element in elements if target_email.lower() in element.text.lower()]
        match_count = len(matching_elements)

        if match_count == 0:
            print(f"No matching email found for '{target_email}'.")
            return False
        elif match_count > 1:
            print(f"Duplicate detected: More than one field contains '{target_email}'!")

        # Click the first matching element
        print(f"Clicking on: {matching_elements[0].text}")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/div/div[2]/form/fieldset/div[2]/div/div/div/ul/li"))
        ).click()
        return True

    except Exception as e:
        print(f"Error occurred: {e}")
        return False


def check_for_duplicate_name(driver, target_name):
    # Find all elements that may contain the name
    elements = driver.find_elements(By.XPATH, "/html/body/div[6]/div[2]/div/div[2]/form/fieldset/div[2]/div/div/div/ul/li")

    # Extract the text from each element and check for occurrences
    match_count = sum(1 for element in elements if target_name.lower() in element.text.lower())

    # Check if duplicates exist
    if match_count >= 2:
        print(f"Duplicate detected: More than one field contains '{target_name}'!")
        return True
    else:
        print(f"No duplicates found for '{target_name}'.")
        return False



def copy_assignments(driver,Link, Name,Email):
    driver.get(Link)
    driver.maximize_window()
    driver.refresh()
    t.sleep(10)
    Copy_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div[3]/div[3]/div/div/div[1]/div[14]'))
    )
    Copy_button.click()
    t.sleep(5)
    select_all_assignments(driver)
    input_field = driver.find_element(By.XPATH, '/html/body/div[6]/div[2]/div/div[2]/form/fieldset/div[2]/div/div/ul/li/input')
    input_field.send_keys(f"{Name}")
    t.sleep(3)
    duplicates = check_for_duplicate_name(driver, Name)

    if duplicates == False:
        username_selection = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="CopyToContactIDs_ajmul"]/div/ul/li[1]'))
        )

        # Click the element
        username_selection.click()
    else:
        select_first_matching_email(driver,Email)

    t.sleep(4)
    # print("test")

    confirm_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[6]/div[2]/div/div[3]/input[1]"))
    )   
    confirm_button.click()
    # print("test3")

def delete_assignments(driver,Link, Name,unassigned_users):
    driver.get(Link)
    driver.maximize_window()
    driver.refresh()

    t.sleep(7)

    Delete_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="actionbar"]/div/div[1]/div[13]'))
    )
    Delete_button.click()

    t.sleep(3)

    # this is the confirm button \/
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div[2]/div/div[3]/input[1]'))
    ).click()

    t.sleep(10)


    # keep a list of user's that we have delete assignments for. we dont need unaccounted for user's
    unassigned_users.loc[len(unassigned_users)] = [Name]



