# uncomment driver.quit()

from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import os
import shutil
import subprocess

import io
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

import pythoncom
import time

import fitz  # PyMuPDF

import win32com.client as win32
import requests
import wget
import zipfile
import json
import platform
import pandas as pd


# ──────────────────────────────────────────────
# Helper Utilities
# ──────────────────────────────────────────────

def get_last_name(name):
    return name.split()[-1]

def get_first_name(name):
    last_name = get_last_name(name)
    last_name_len = len(last_name)
    first_name = name[:-(last_name_len + 1)]
    return first_name


def check_system_bit():
    bits = platform.architecture()[0]
    return bits

global Bits
Bits = check_system_bit()
Bits = Bits.split("b")[0]


def ShowNameAndTokenNo(token_no, Name):
    js = """
        try{
            document.getElementById('my_box').remove();
        }
        catch (error) {
            console.log('my box ');
        }
        var div = document.createElement('div');
        div.setAttribute('id', 'my_box')
        div.innerText = 'Name: """ + Name + """; Token_No: """ + token_no + """'
        div.style.fontSize = "xx-large";
        div.style.color = "white";
        div.style.margin = "0.5% 0";
        div.style.backgroundColor = "red";
        div.style.ZIndex = "100";
        div.style.position = 'fixed';
        div.style.top = 0;
        div.style.left = "30vh";

        document.body.appendChild(div);
        """
    return js


# ──────────────────────────────────────────────
# Credential Loaders
# ──────────────────────────────────────────────

def ApnaPanId():
    with open('apna_pan_id.txt') as file:
        return file.read().strip()

def ApnaPanPassword():
    with open('apna_pan_password.txt') as file:
        return file.read().strip()

def NsdlID():
    with open('nsdl_id.txt') as file:
        return file.read().strip()

def NsdlPassword():
    with open('nsdl_password.txt') as file:
        return file.read().strip()


# ──────────────────────────────────────────────
# Driver Management
# ──────────────────────────────────────────────

def asign_browser():
    try:
        browser_path = get_driver_path()
    except Exception:
        browser_path = download_browser()
    return browser_path


def download_browser():
    def get_version_via_com(filename):
        parser = win32.Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(filename)
        except Exception:
            return None
        return version

    browser_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]

    version = list(filter(None, [get_version_via_com(p) for p in browser_paths]))[0]
    version_list = version.split(".")
    version_txt = f"{version_list[0]}.{version_list[1]}.{version_list[2]}"
    folder_name = version_txt

    driver_path = (
        f'{os.getcwd()}\\drivers\\chromedriver\\win32\\{folder_name}'
        f'\\chromedriver-win{Bits}\\chromedriver.exe'
    )

    get_driver_download_version = (
        'https://googlechromelabs.github.io/chrome-for-testing/'
        'known-good-versions-with-downloads.json'
    )

    response = requests.get(get_driver_download_version)
    response_data = response.json()
    chrome_versions = response_data['versions']

    download_url = None
    for chrome_version in chrome_versions:
        if version_txt in chrome_version['version']:
            idx = 4 if Bits == '64' else 3
            download_url = chrome_version['downloads']['chromedriver'][idx]['url']
            break

    if not download_url:
        raise RuntimeError(f"Could not find chromedriver for Chrome version {version_txt}")

    latest_driver_zip = wget.download(download_url, f'{os.getcwd()}\\chromedriver.zip')
    print('\n')

    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall(
            path=f'{os.getcwd()}\\drivers\\chromedriver\\win32\\{folder_name}'
        )

    os.remove(latest_driver_zip)

    # Persist path
    try:
        with open('driver_location.json') as file:
            value = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        value = {}

    value['chrome_path'] = driver_path
    with open('driver_location.json', 'w') as file:
        json.dump(value, file)

    return driver_path


def get_driver_path():
    with open('driver_location.json') as file:
        value = json.load(file)
    return value['chrome_path']


def MinimizeWindow():
    try:
        driver.minimize_window()
    except Exception:
        pass

def MaximizeWindow():
    try:
        driver.maximize_window()
    except Exception:
        pass


def InitiliseBrowser():
    global driver
    pythoncom.CoInitialize()

    options = Options()
    download_location = os.getcwd() + "\\Downloads"
    
    os.makedirs(download_location, exist_ok=True)

    prefs = {
        "download.default_directory": download_location,
        "safebrowsing.enabled": "false"
    }
    options.add_experimental_option("prefs", prefs)

    try:
        service = Service(executable_path=asign_browser())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception:
        # Fall back to webdriver_manager auto-install
        service = Service(ChromeDriverManager().install())

        # Persist the path for next time
        driver_path = service.path
        wdm_json_path = f'{os.getcwd()}\\driver\\.wdm\\driver_location.json'
        try:
            os.makedirs(os.path.dirname(wdm_json_path), exist_ok=True)
            try:
                with open(wdm_json_path) as f:
                    driver_paths = json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                driver_paths = {}
            driver_paths['chrome_path'] = driver_path
            with open(wdm_json_path, 'w') as f:
                json.dump(driver_paths, f)
        except Exception:
            pass

        driver = webdriver.Chrome(service=service, options=options)

    driver.get(f"file:///{os.getcwd()}/version.txt")
    MinimizeWindow()


# ──────────────────────────────────────────────
# Login Functions
# ──────────────────────────────────────────────

def LoginApnaPan():
    driver.get('https://apnapanindia.co.in/soft/SuperAdmin/op-login.php')
    driver.maximize_window()

    driver.find_element(By.ID, 'email').send_keys(ApnaPanId())
    driver.find_element(By.ID, 'Password').send_keys(ApnaPanPassword())
    driver.find_elements(By.TAG_NAME, 'button')[0].click()
    return "Login Successfully"


def LoginNSDL():
    try:
        driver.execute_script("window.open('https://www.onlineservices.proteantech.in/paam/', 'newwindow')")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(3)

        # Accept cookie banner
        try:
            driver.find_element(By.ID, 'onetrust-accept-btn-handler').click()
        except Exception:
            pass

        # Dismiss any browser alert
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass

        def _fill_and_submit():
            driver.find_element(By.ID, 'userID').send_keys(NsdlID())
            driver.find_element(By.ID, 'password').send_keys(NsdlPassword())
            while True:
                driver.find_element(By.ID, 'captcha').click()
                cap_len = len(driver.find_element(By.ID, 'captcha').get_attribute("value"))
                if cap_len == 5:
                    driver.find_element(By.ID, 'submit_btn').click()
                    break

        _fill_and_submit()

        try:
            time.sleep(3)
            error = driver.find_element(By.ID, 'jsError').text

            if error == 'Invalid Captcha':
                while True:
                    _fill_and_submit()
                    try:
                        error = driver.find_element(By.ID, 'jsError').text
                        if error != "Invalid Captcha":
                            break
                    except Exception:
                        break
            else:
                if error == "":
                    time.sleep(5)
                else:
                    command = (
                        f'mshta vbscript:Execute("CreateObject(""WScript.Shell"")'
                        f'.Popup ""{error}"", 0, ""Information"":close")'
                    )
                    subprocess.Popen(command)
                    driver.quit()
                    return False

        except Exception as e:
            print(e)

        time.sleep(2)
        if driver.current_url == 'https://www.onlineservices.proteantech.in/paam/login.html':
            time.sleep(0.25)
            try:
                driver.execute_script("closeDialog()")
            except Exception:
                pass
            try:
                driver.find_element(By.ID, 'emailId').click()
                time.sleep(1)
                driver.find_element(By.ID, 'submit_btn').click()
            except Exception as e:
                print(e)

        while True:
            time.sleep(2)
            url = driver.current_url
            if "https://www.onlineservices.proteantech.in/paam/validateOTPTinfc.html?ID" in url:
                time.sleep(2)
                if "https://www.onlineservices.proteantech.in/paam/validateOTPTinfc.html?ID" in driver.current_url:
                    break

        return True

    except Exception as e:
        # command = (
        #     f'mshta vbscript:Execute("CreateObject(""WScript.Shell"")'
        #     f'.Popup ""{e}"", 0, ""Information"":close")'
        # )
        # subprocess.Popen(command)
        # return False
        LoginNSDL()
# ──────────────────────────────────────────────
# NSDL Form Submission – New PAN
# ──────────────────────────────────────────────

def NewPersonNSDL(all_details):
    try:
        driver.switch_to.window(driver.window_handles[1])

        try:
            driver.find_element(By.ID, 'oCMenu_top3').click()
            time.sleep(0.25)
            driver.find_element(By.ID, 'oCMenu_sub10').click()
            time.sleep(0.25)
        except Exception:
            return {"verify_status": 'not done', "message": 'SESSION EXPIRE'}

        try:
            driver.execute_script(
                ShowNameAndTokenNo(
                    token_no=all_details['PanCard Serial ID'],
                    Name=all_details['Name']
                )
            )
        except Exception:
            pass

        # Category
        Select(driver.find_element(By.ID, 'cat_applicant')).select_by_value(
            all_details['Category of Applicant']
        )

        # Split name
        user_name = all_details['Name']

        l_name = get_last_name(user_name)
        f_name = get_first_name(user_name)

        driver.find_element(By.ID, 'l_name').send_keys(l_name)

        # Name on card via JS (safer for special characters)
        # name_on_aadhar = all_details['Name on Card']
        # try:
        #     driver.execute_script(
        #         f"document.getElementById('name_card').value = '{name_on_aadhar}';"
        #     )
        # except Exception:
        #     driver.find_element(By.ID, 'name_card').send_keys(name_on_aadhar)


        # remove the try here
        driver.find_element(By.ID, 'f_name').send_keys(f_name)


        driver.find_element(By.ID, 'nmother').click()

        # Father name
        father_name = all_details["Father's Name"]
        driver.find_element(By.ID, 'fal_name').send_keys(get_last_name(father_name))
        driver.find_element(By.ID, 'faf_name').send_keys(get_first_name(father_name))

        # Mother name
        mother_name = all_details["Mother's Name"]
        driver.find_element(By.ID, 'opaMotherFirstName').send_keys(get_first_name(mother_name))
        driver.find_element(By.ID, 'opaMotherLastName').send_keys(get_last_name(mother_name))

        # Date of birth
        birth_date = str(all_details['Date of Birth']).replace('-', '/')

        time.sleep(2)
        driver.find_element(By.ID, 'dob').send_keys(birth_date)

        # residential_status
        time.sleep(0.25)
        Select(driver.find_element(By.ID, 'residential_status')).select_by_value('R')

        # Phone
        Select(driver.find_element(By.ID, 'mobile_isd')).select_by_value('91')

        driver.find_element(By.ID, 'mobile_number').send_keys(all_details['Contact No'])
        try:
            driver.switch_to.alert.accept()
            time.sleep(1)
        except Exception:
            pass

        # driver.find_element(By.ID, 'tel_num_stdcode').send_keys("")
        # driver.find_element(By.ID, 'tel_num').send_keys(all_details['Contact No'])

        # Email & Address
        driver.find_element(By.ID, 'email_id').send_keys(all_details['Email Id'])
        driver.find_element(By.ID, 'add_comm').send_keys('INDIAN')
        # driver.find_element(By.ID, 'ra_add').click()

        try:
            time.sleep(1)
            driver.switch_to.alert.accept()
        except Exception:
            pass

        # driver.find_element(By.ID, 'ra_add').send_keys('INDIAN')

        # Proof documents
        for field_id in ('proof_id', 'proof_add'):
            Select(driver.find_element(By.ID, field_id)).select_by_visible_text(
                "AADHAAR Card issued by the Unique Identification Authority of India"
            )
        
        proof_of_dob = all_details['Proof of DOB'].strip()
        if proof_of_dob == "Matriculation certificate or mark sheet of recognized board":
            proof_of_dob = "Matriculation Marksheet of recognised board"

        elif proof_of_dob == "Electors Photo Identity Card (VOTER CARD )":
            proof_of_dob = "Elector's photo identity card"

        try:
            time.sleep(1)
            select_el = Select(driver.find_element(By.ID, 'proof_dob'))
            
            # Try matching by visible text first (stripped)
            matched = False
            for option in select_el.options:

                if option.text.strip() == proof_of_dob:
                    option.click()
                    matched = True
                    break

            # Fallback: select by index if text starts with same words
            if not matched:
                for option in select_el.options:
                    if option.text.strip().startswith(proof_of_dob[:30]):
                        option.click()
                        matched = True
                        break
            
            # Last resort: use JavaScript to set by value/text
            if not matched:
                driver.execute_script("""
                    var select = document.getElementById('proof_dob');
                    for (var i = 0; i < select.options.length; i++) {
                        if (select.options[i].text.trim().includes(arguments[0])) {
                            select.selectedIndex = i;
                            select.dispatchEvent(new Event('change'));
                            break;
                        }
                    }
                """, proof_of_dob[:50])

        except Exception as e:
            print(f"proof_dob selection error: {e}")

        # Assessing office
        assessing_office = all_details['Assessing Office'].replace(" ", "").split("|")
        if any(v == "" for v in assessing_office[:4]):
            return {
                "verify_status": "not done",
                "message": (
                    "nsdl Id me error - Area Code is mandatory "
                    "(Apna Pan india me Ao Code Hi mandatory kra deta hu)"
                )
            }

        driver.find_element(By.ID, 'area_code').send_keys(assessing_office[0])
        driver.find_element(By.ID, 'ao_type').send_keys(assessing_office[1])
        driver.find_element(By.ID, 'range_code').send_keys(assessing_office[2])
        driver.find_element(By.ID, 'ao_num').send_keys(assessing_office[3])

        state = all_details['Pan Card Dispatched State']

        driver.find_element(By.ID, 'aoSelection').click()
        Select(driver.find_element(By.ID, 'state_aoCode')).select_by_visible_text(state)
        Select(driver.find_element(By.ID, 'user_state')).select_by_visible_text(str(state))

        # Verify state selection took effect
        state_value = driver.find_element(By.ID, 'user_state').get_attribute('value')
        if state_value == "Please Select":
            time.sleep(1)
            Select(driver.find_element(By.ID, 'user_state')).select_by_visible_text(str(state))


        # Enter the pincode
        driver.find_element(By.ID, 'pincode').click()
        driver.find_element(By.ID, 'pincode').send_keys(all_details['Pincode'])


        # Aadhaar
        Select(driver.find_element(By.ID, 'check_aadhaar_eid')).select_by_value('A')
        driver.find_element(By.ID, 'aadhaarNo').send_keys(all_details['Aadhar No'])
        driver.find_element(By.ID, 'name_aadhaar').send_keys(all_details['Name on Card'])
        Select(driver.find_element(By.ID, 'gender')).select_by_value(all_details['Gender'])
        driver.find_element(By.ID, 'aadhaarconsent').click()
        driver.find_element(By.ID, 'verify').click()

        time.sleep(2.5)
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        time.sleep(2.5)

        # Poll result
        verify_status = "not done"
        return_message = ""
        for n in range(10):
            time.sleep(2.5)
            return_message = driver.find_element(By.ID, 'result_message').text

            if return_message == 'Demographic verification successful':
                verify_status = "done"
                try:
                    driver.switch_to.alert.accept()
                except Exception:
                    pass
                time.sleep(1)
                driver.find_element(By.ID, 'proceed').click()
                try:
                    driver.switch_to.alert.accept()
                except Exception:
                    pass
                time.sleep(1)
                driver.find_element(By.ID, 'submit').click()
                time.sleep(5)
                break

            elif return_message in (
                'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.',
                'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.',
                'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in',
            ):
                verify_status = "not done"
                break

            elif 'Alert Text: Area Code is mandatory' in return_message:
                verify_status = "not done"
                break

            elif return_message == "":
                time.sleep(2.5)
            else:
                verify_status = "not done"
                break

    except Exception as e:
        verify_status = "not done"
        return_message = str(e)
        print(return_message)

    return_data = {"verify_status": verify_status, "message": return_message}
    print('\n\n', return_data, '\n\n')
    return return_data



# ──────────────────────────────────────────────
# NSDL Form Submission – PAN Correction
# ──────────────────────────────────────────────

def CorrectionPersonNSDL(all_details):
    try:
        driver.switch_to.window(driver.window_handles[1])

        try:
            driver.find_element(By.ID, 'oCMenu_top3').click()
            time.sleep(0.25)
        except Exception:
            pass

        try:
            driver.find_element(By.ID, 'oCMenu_sub12').click()
            time.sleep(0.25)
        except Exception:
            return {"verify_status": 'not done', "message": 'SESSION EXPIRE'}

        try:
            driver.execute_script(
                ShowNameAndTokenNo(
                    token_no=all_details['PanCard Serial ID'],
                    Name=all_details['Name']
                )
            )
        except Exception:
            pass

        # PAN number – use JS as a reliable fallback
        pan_el = driver.find_element(By.ID, 'pan')
        pan_el.click()
        pan_el.send_keys(all_details['Old PanCard No'])
        if not pan_el.get_attribute('value'):
            driver.execute_script(
                f"document.getElementById('pan').value = '{all_details['Old PanCard No']}';"
            )

        Select(driver.find_element(By.ID, 'category')).select_by_value(
            all_details['Category of Applicant']
        )

        # Name
        name_list = all_details['Name'].split(" ")
        l_name = name_list[-1]
        f_name = all_details['Name'].replace(l_name, "").strip()

        driver.find_element(By.ID, 'lastName').send_keys(l_name)
        driver.find_element(By.ID, 'firstName').send_keys(f_name)

        # Father name
        father_name = all_details["Father's Name"]
        driver.find_element(By.ID, 'fatherlastName').send_keys(get_last_name(father_name))
        driver.find_element(By.ID, 'fatherfirstName').send_keys(get_first_name(father_name))

        # it should be click before adding mother name here
        driver.find_element(By.ID, 'crSubmitBtn').click()

        try:
            time.sleep(2)
            driver.switch_to.alert.accept()
        except Exception:
            pass


        # Mother name
        mother_name = all_details["Mother's Name"]
        
        driver.find_element(By.ID, 'opaMotherLastName').send_keys(get_last_name(mother_name))
        driver.find_element(By.ID, 'opaMotherFirstName').send_keys(get_first_name(mother_name))

        driver.find_element(By.ID, 'crSubmitBtn').click()



        try:
            time.sleep(2)
            driver.switch_to.alert.accept()
        except Exception:
            pass

        driver.find_element(By.ID, 'acceptSubmit').click()
        time.sleep(1)
        # driver.find_element(By.ID, 'nameOnCard').send_keys(all_details['Name on Card'])

        # DOB
        birth_date = str(all_details['Date of Birth']).replace('-', '/')
        time.sleep(2)
        driver.find_element(By.ID, 'dob').send_keys(birth_date)

        # ISD Details
        Select(driver.find_element(By.ID, 'mobile_isd')).select_by_value('91')
        driver.find_element(By.ID, 'mobile_number').send_keys(all_details["Contact No"])

        Select(driver.find_element(By.ID, 'add_comm')).select_by_value('indian')
        driver.find_element(By.ID, 'emailId').send_keys(all_details['Email Id'])

        state = all_details['Pan Card Dispatched State']
        Select(driver.find_element(By.ID, 'user_state')).select_by_visible_text(str(state))

        Select(driver.find_element(By.ID, 'check_aadhaar_eid')).select_by_value('A')
        driver.find_element(By.ID, 'aadhaarNo').send_keys(all_details['Aadhar No'])
        driver.find_element(By.ID, 'name_aadhaar').send_keys(all_details['Name on Card'])

        for field_id in ('poid', 'por'):
            Select(driver.find_element(By.ID, field_id)).select_by_visible_text(
                "AADHAAR Card issued by the Unique Identification Authority of India"
            )
        
        proof_of_dob = all_details['Proof of DOB'].strip()
        if proof_of_dob == "Matriculation certificate or mark sheet of recognized board":
            proof_of_dob = "Matriculation Marksheet of recognised board"

        elif proof_of_dob == "Electors Photo Identity Card (VOTER CARD )":
            proof_of_dob = "Elector's photo identity card"

        try:
            time.sleep(1)
            select_el = Select(driver.find_element(By.ID, 'pod'))
            
            # Try matching by visible text first (stripped)
            matched = False
            for option in select_el.options:

                if option.text.strip() == proof_of_dob:
                    option.click()
                    matched = True
                    break

            # Fallback: select by index if text starts with same words
            if not matched:
                for option in select_el.options:
                    if option.text.strip().startswith(proof_of_dob[:30]):
                        option.click()
                        matched = True
                        break
            
            # Last resort: use JavaScript to set by value/text
            if not matched:
                driver.execute_script("""
                    var select = document.getElementById('pod');
                    for (var i = 0; i < select.options.length; i++) {
                        if (select.options[i].text.trim().includes(arguments[0])) {
                            select.selectedIndex = i;
                            select.dispatchEvent(new Event('change'));
                            break;
                        }
                    }
                """, proof_of_dob[:50])

        except Exception as e:
            print(f"proof_dob selection error: {e}")

        # Enter the pincode
        driver.find_element(By.ID, 'pincode').click()
        driver.find_element(By.ID, 'pincode').send_keys(all_details['Pincode'])


        Select(driver.find_element(By.ID, 'gender')).select_by_value(all_details['Gender'])
        driver.find_element(By.ID, 'aadhaarconsent').click()
        driver.find_element(By.ID, 'verify').click()

        time.sleep(2.5)

        verify_status = "not done"
        return_message = ""
        while True:
            return_message = driver.find_element(By.ID, 'result_message').text

            if return_message == 'Demographic verification successful.':
                verify_status = "done"
                driver.find_element(By.ID, 'Save').click()
                time.sleep(0.25)
                driver.find_element(By.ID, 'SaveConfirm').click()
                time.sleep(2.5)
                break

            elif return_message in (
                'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.',
                'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.',
                'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in',
            ):
                verify_status = "not done"
                break

            elif 'Alert Text: Area Code is mandatory' in return_message:
                verify_status = "not done"
                break

            elif return_message == "":
                time.sleep(5)
            else:
                verify_status = "not done"
                break

    except Exception as e:
        verify_status = "not done"
        return_message = str(e)


    return_data = {"verify_status": verify_status, "message": return_message}
    print('\n\n', return_data, '\n\n')
    return return_data


# ──────────────────────────────────────────────
# File / PDF Utilities
# ──────────────────────────────────────────────

def get_latest_file_in_folder(dir_path, file_type):
    files = [f for f in os.listdir(dir_path) if f.endswith(file_type)]
    file_info = [(f, os.stat(os.path.join(dir_path, f)).st_mtime) for f in files]
    file_info.sort(key=lambda x: x[1])
    return file_info[-1][0]


def CheckForNewFile(old_pdf_file):
    download_location = os.getcwd() + '\\Downloads'

    
    while True:
        time.sleep(0.5)
        new_pdf_file = get_latest_file_in_folder(download_location, ".pdf")
        if new_pdf_file != old_pdf_file:
            break
    return f"{download_location}\\{new_pdf_file}"


def GeneratePNumber(file_name):
    pdf_document = fitz.open(file_name)
    text = ""
    for page_num in range(pdf_document.page_count):
        text = pdf_document[page_num].get_text()
    pdf_document.close()

    text_list = text.split('Proof of DOB')
    text_list = text_list[1].split('Date')
    text = text_list[0].replace("\n", "").replace(" ", "").replace("P-", "")

    new_name = os.getcwd() + "\\Downloads\\" + text + "hhhh.pdf"
    os.rename(file_name, new_name)
    return text


# ──────────────────────────────────────────────
# Result Entry Back in ApnaPan
# ──────────────────────────────────────────────

def EnterNSDLResult(status, message):
    STATUS_MESSAGES = {
        'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.':
            'Please Wait Sometime for acknowledgement! Slip',
        'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.':
            'Pan has already issued Been Aadhar (Pan Pahle se Ban gya hai is Aadhar Se)'
            'Find To Pan Box No - 04 se Search karen ',
        'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in':
            'Online Aap Aadhar Number Wrong Fill kiye hai... '
            '(Check - Online Aadhar Number/Name/DoB/Sppeling)..Wait For Ack Slip',
    }

    if not status:
        friendly_message = STATUS_MESSAGES.get(message, None)

        if friendly_message is None:
            if 'Alert Text: Area Code is mandatory' in message:
                friendly_message = 'Next Time Online AO Code Fill Kiya Karen,Wait karen Ack Slip ke liye.'
            else:
                friendly_message = "Please Wait Sometime for acknowledgement! Slip"

        if message == list(STATUS_MESSAGES.keys())[2]:   # demographic verification error
            friendly_message += " Please Wait Sometime for acknowledgement! Slip"

        time.sleep(1)
        driver.switch_to.window(driver.window_handles[0])
        Select(driver.find_element(By.ID, 'Status')).select_by_value("5")
        driver.find_element(By.NAME, 'Remarks').clear()
        driver.find_element(By.NAME, 'Remarks').send_keys(friendly_message)
        driver.find_elements(By.NAME, 'submit')[0].click()

    else:
        download_location = os.getcwd() + "\\Downloads"
        

        print(download_location)
        old_pdf_file = get_latest_file_in_folder(download_location, ".pdf")
        print(old_pdf_file)

        try:
            pyautogui.moveTo(800, 500)
            pyautogui.rightClick()
            time.sleep(0.25)
            pyautogui.keyDown('down')
            time.sleep(0.25)
            pyautogui.keyDown('Enter')
            time.sleep(1.5)
            pyautogui.keyDown('Enter')
        except Exception:
            pass

        try:
            time.sleep(2)
            file_name = CheckForNewFile(old_pdf_file)
            p_number = GeneratePNumber(file_name)
            print('P Number: ', p_number)

            print('Going to click on home button')
            driver.find_elements(By.TAG_NAME, 'input')[0].click()
            driver.switch_to.window(driver.window_handles[0])

            Select(driver.find_element(By.ID, 'Status')).select_by_value("1")
            time.sleep(0.5)

            file_name = os.getcwd() + '\\Downloads\\' + p_number + "hhhh.pdf"
            driver.find_element(By.NAME, 'ACkPdf').send_keys(file_name)
            driver.find_element(By.NAME, 'ACKNo').send_keys(p_number)
            driver.find_element(By.NAME, 'Remarks').clear()
            driver.find_element(By.NAME, 'Remarks').send_keys(".")
            driver.find_elements(By.NAME, 'submit')[0].click()

            time.sleep(0.25)
            driver.execute_script("document.getElementsByClassName('btn-final-pdf')[0].click()")

        except Exception as e:
            print('\n\n error => ', e, '\n\n')
            driver.switch_to.window(driver.window_handles[1])
            command = (
                'mshta vbscript:Execute("CreateObject(""WScript.Shell"")'
                '.Popup ""Please Fill This PDF Manually, and After Filling Click on Home Button"","'
                ' 0, ""Information"":close")'
            )
            subprocess.Popen(command)
            time.sleep(2)

            while True:
                time.sleep(1)
                if "onlineservices.proteantech.in/paam/homeTinFC.html?ID" in driver.current_url:
                    print('i start my work again')
                    time.sleep(3)
                    driver.switch_to.window(driver.window_handles[0])
                    break



# ──────────────────────────────────────────────
# Main Entry Point
# ──────────────────────────────────────────────

InitiliseBrowser()
apna_pan_login_status = LoginApnaPan()

g = LoginNSDL()

# g = True


global removed_states
removed_states = ['ASSAM', 'MIZORAM', 'JAMMU AND KASHMIR']

if g:
    while True:
        try:
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            if "apnapanindia.co.in" not in driver.current_url:
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[0])
        
  
        driver.get('https://apnapanindia.co.in/soft/SuperAdmin/pancard-power/op-index.php')

        Select(driver.find_elements(By.NAME, 'PanType')[0]).select_by_value('Individual')
        Select(driver.find_elements(By.NAME, 'Status')[0]).select_by_value('4')
        driver.find_elements(By.CLASS_NAME, 'filtr')[0].click()

        try:
            setPage = driver.find_elements(By.CLASS_NAME, 'setPage')[0].text
            setPage = setPage.replace(" ", "").split("of")[1]
            driver.get(
                f'https://apnapanindia.co.in/soft/SuperAdmin/pancard-power/op-index.php?page={setPage}'
            )
        except Exception:
            pass

        btns = driver.find_elements(By.CLASS_NAME, 'DetailBTN')

        all_urls = [btn.get_attribute('href') for btn in btns]

        if not all_urls:
            command = (
                'mshta vbscript:Execute("CreateObject(""WScript.Shell"")'
                '.Popup ""All work is completed..."", 0, ""Information"":close")'
            )
            subprocess.Popen(command)
            driver.quit()
            exit()
        
        for url in all_urls:
            driver.get(url)

            html_content = driver.find_elements(By.TAG_NAME, 'table')[0].get_attribute('outerHTML')
            df = pd.read_html(io.StringIO(html_content))[0]

            df = df.to_dict(orient='records')

            page_data = {row[0]: str(row[1]) for row in df}

            print(page_data)

            if page_data['Pan Card Dispatched State'] in removed_states:
                work_status = {
                    'verify_status': 'not done',
                    'message': "Please wait for Acknowledgement slip in PAN card."
                }
            
            elif page_data['PanCard Type'] == 'New':
                for _ in range(3):
                    work_status = NewPersonNSDL(page_data)
                    if work_status['verify_status'] == "not done":
                        print(work_status['message'])
                        if work_status['message'] == 'SESSION EXPIRE':
                            command = (
                                'mshta vbscript:Execute("CreateObject(""WScript.Shell"")'
                                '.Popup ""Your Session is Expired..."", 0, ""Information"":close")'
                            )
                            subprocess.Popen(command)
                            driver.quit()
                            exit()
                        # Break on known terminal errors
                        if work_status['message'] in (
                            'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.',
                            'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.',
                            'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in',
                        ) or 'Alert Text: Area Code is mandatory' in work_status['message']:
                            break
                        try:
                            driver.switch_to.alert.accept()
                        except Exception:
                            pass
                    else:
                        break

            elif page_data['PanCard Type'] == 'Correction':

                for _ in range(3):
                    work_status = CorrectionPersonNSDL(page_data)
                    if work_status['verify_status'] == "not done":
                        if work_status['message'] == 'SESSION EXPIRE':
                            command = (
                                'mshta vbscript:Execute("CreateObject(""WScript.Shell"")'
                                '.Popup ""Your Session is Expired..."", 0, ""Information"":close")'
                            )
                            subprocess.Popen(command)
                            driver.quit()
                            exit()
                        if work_status['message'] in (
                            'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.',
                            'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.',
                            'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in',
                        ) or 'Alert Text: Area Code is mandatory' in work_status['message']:
                            break
                        try:
                            driver.switch_to.alert.accept()
                        except Exception:
                            pass
                    else:
                        break
                
            if work_status['verify_status'] == "done":
                EnterNSDLResult(True, ".")
            elif work_status['verify_status'] == 'not done':
                EnterNSDLResult(False, work_status['message'])
