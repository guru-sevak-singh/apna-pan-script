from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import os
import shutil
import subprocess

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

import pythoncom
import time

import fitz  

from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options

import win32com.client as win32
import requests
import wget
import zipfile
import json
import platform
import pandas as pd


def get_last_name(name):
    return name.split()[-1]

def get_first_name(name):
    last_name = get_last_name(name)
    last_name_len = len(last_name)
    first_name = name[:-(last_name_len+1)]
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

def ApnaPanId():
    with open('apna_pan_id.txt') as file:
        id = file.read()
    return id

def ApnaPanPassword():
    with open('apna_pan_password.txt') as file:
        password = file.read()
    return password

def NsdlID():
    with open('nsdl_id.txt') as file:
        id = file.read()
    return id

def NsdlPassword():
    with open('nsdl_password.txt') as file:
        password = file.read()
    return password


def asign_browser():
    try:
        browser_path = get_driver_path()
    except:
        browser_path = download_borwser()
    
    return browser_path

def download_borwser():

    def get_version_via_com(filename):
        parser = win32.Dispatch("Scripting.FileSystemObject")
        try:
            version = parser.GetFileVersion(filename)
        except Exception:
            return None
        return version
    

    browser_paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    
    version = list(filter(None, [get_version_via_com(p) for p in browser_paths]))[0]

    version_list = version.split(".")
    
    version_txt = f"{version_list[0]}.{version_list[1]}.{version_list[2]}"

    folder_name = version_txt

    driver_path = f'{os.getcwd()}\drivers\chromedriver\win32\{folder_name}\chromedriver-win{Bits}\chromedriver.exe'

    get_driver_download_version = 'https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json'

    response = requests.get(get_driver_download_version)
    response_data = response.json()
    
    
    chrome_versions = response_data['versions']
    
    for chrome_version in chrome_versions:
        
        if version_txt in  chrome_version['version']:
            if Bits == '64':
                download_url = chrome_version['downloads']['chromedriver'][4]['url']
            else:
                download_url =  chrome_version['downloads']['chromedriver'][3]['url']
            break
        else:
            continue   
    
    latest_driver_zip = wget.download(download_url,f'{os.getcwd()}\chromedriver.zip')
    
    print('\n')
    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        
        zip_ref.extractall(path=f'{os.getcwd()}\drivers\chromedriver\win32\{folder_name}') 

    os.remove(latest_driver_zip)

    try:
        with open('dop_driver.json') as file:
            value = json.load(file)
        with open('dop_driver.json','w') as file:
            value['chrome_path'] = driver_path
            json.dump(value, file)
    except:
        with open('dop_driver.json', 'w') as file:
            value = {'chrome_path': driver_path}
            json.dump(value, file)

    return driver_path
    


def get_driver_path():
    
    with open('dop_driver.json') as file:
        value = json.load(file)
    path = value['chrome_path']
    
    return path

def MinimizeWindow():
    try:
        driver.minimize_window()
    except:
        pass

def MaxemizeWindow():
    try:
        driver.maximize_window()
    except:
        pass


def InitiliseBrowser():
    global driver
    pythoncom.CoInitialize()

    options = Options()

    download_location = os.getcwd()+"\\Downloads"

    prefs={"download.default_directory": download_location,"safebrowsing.enabled": "false"}
    options.add_experimental_option("prefs",prefs)

    
    try:
        driver = webdriver.Chrome(executable_path=asign_browser(), options=options)
    except:
        driver_path = ChromeDriverManager().install()
        try:
            with open(f'{os.getcwd()}\\driver\\.wdm\dop_driver.json') as file:
                driver_paths = json.load(file)

            with open(f'{os.getcwd()}\\driver\\.wdm\dop_driver.json', 'w') as file:
                driver_paths['chrome_path'] = driver_path
                json.dump(driver_paths, file)
        except FileNotFoundError:
            driver_paths = {}
            with open(f'{os.getcwd()}\\driver\\.wdm\dop_driver.json', 'w') as file:
                driver_paths['chrome_path'] = driver_path
                json.dump(driver_paths, file)

        driver = webdriver.Chrome(executable_path = driver_path, options=options)

    driver.get(f"{os.getcwd()}//version.txt")
    MinimizeWindow()




def LoginApnaPan():
    driver.get('https://apnapanindia.co.in/soft/SuperAdmin/op-login.php')
    driver.maximize_window()
    email_box = driver.find_element_by_id('email')
    password_box = driver.find_element_by_id('Password')
    email_box.send_keys(ApnaPanId())
    password_box.send_keys(ApnaPanPassword())
    submit_button = driver.find_elements_by_tag_name('button')[0]
    submit_button.click()
    return "Login Successfully"

def LoginNSDL():
    try:
        driver.execute_script(f"window.open('https://www.onlineservices.nsdl.com/paam/', 'newwindow')")
        driver.switch_to.window(driver.window_handles[1])

        
        time.sleep(3)
        try:
            driver.find_element_by_id('onetrust-accept-btn-handler').click()
        except:
            pass

        try:
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass
        
        user_id = driver.find_element_by_id('userID')
        user_id.send_keys(NsdlID())
        password = driver.find_element_by_id('password')
        password.send_keys(NsdlPassword())
        while True:

            driver.find_element_by_id('captcha').click()
            cap_box = len(driver.find_element_by_id('captcha').get_attribute("value"))
            print(cap_box)
            if cap_box == 5:
                driver.find_element_by_id('submit_btn').click()
                break
            else:
                continue
        
        try:
            time.sleep(3)
            error = driver.find_element_by_id('jsError').text

            if error == 'Invalid Captcha':
                while True:
                    user_id = driver.find_element_by_id('userID')
                    user_id.send_keys(NsdlID())
                    password = driver.find_element_by_id('password')
                    password.send_keys(NsdlPassword())
                    while True:
                        driver.find_element_by_id('captcha').click()
                        cap_box = len(driver.find_element_by_id('captcha').get_attribute("value"))
                        if cap_box == 5:
                            driver.find_element_by_id('submit_btn').click()
                            break
                        else:
                            continue
                    try:

                        error = driver.find_element_by_id('jsError').text
                        if error != "Invalid Captcha":
                            break
                        else:
                            pass
                    except:
                        break

                
                
            else:
                if error == "":
                    time.sleep(5)
                    pass
                else:

                    command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{error}"", 0, ""Information"":close")'
                    subprocess.Popen(command)
                    driver.quit()
                    return False
                
        except Exception as e:
            print(e)


        time.sleep(2)
        if driver.current_url == 'https://www.onlineservices.nsdl.com/paam/login.html':
            time.sleep(0.25)
            try:
                js = "closeDialog()"
                driver.execute_script(js)

            except:
                pass
            try:
                driver.find_element_by_id('emailId').click()
                time.sleep(1)
                driver.find_element_by_id('submit_btn').click()
            except Exception as e:
                print(e)
        

        while True:
            time.sleep(2)
            url = driver.current_url
            if "https://www.onlineservices.nsdl.com/paam/validateOTPTinfc.html?ID" in url:
                time.sleep(2)
                url = driver.current_url
                if "https://www.onlineservices.nsdl.com/paam/validateOTPTinfc.html?ID" in url:
                    break
                else:
                    continue
            else:
                continue
        
        return True

    except Exception as e:
        command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""{e}"", 0, ""Information"":close")'
        subprocess.Popen(command)
        return False

def NewPersonNSDL(all_details):
    try:
        driver.switch_to.window(driver.window_handles[1])
        try:
            driver.find_element_by_id('oCMenu_top3').click()
            time.sleep(0.25)

            driver.find_element_by_id('oCMenu_sub10').click()
            time.sleep(0.25)
        except:
            return_data =  {"verify_status": 'not done', "message": 'SESSION EXPIRE'}
            return return_data

        try:
            js = ShowNameAndTokenNo(token_no=all_details['PanCard Serial ID'], Name=all_details['Name'])
            driver.execute_script(js)
        except:
            pass


        driver.find_element_by_id('cat_applicant').click()
        sel = Select(driver.find_element_by_id('cat_applicant'))
        sel.select_by_value(all_details['Category of Applicant'])


        name_list = all_details['Name'].split(" ")

        l_name = name_list[len(name_list) - 1]
        f_name = all_details['Name'].replace(l_name, "")

        driver.find_element_by_id('l_name').click()
        driver.find_element_by_id('l_name').send_keys(l_name)

        
        name_on_aadhar = all_details['Name on Card']
        try:
            js = f"document.getElementById('name_card').value = '{name_on_aadhar}'; console.log('{name_on_aadhar}');"
            driver.execute_script(js)
        except:
            driver.find_element_by_id('name_card').click()
            driver.find_element_by_id('name_card').send_keys(name_on_aadhar)


        driver.find_element_by_id('f_name').click()
        driver.find_element_by_id('f_name').send_keys(f_name)


        driver.find_element_by_id('nmother').click()


        father_name = all_details["Father's Last Name"]

        # get last name
        father_last_name = get_last_name(father_name)

        driver.find_element_by_id('fal_name').click()
        driver.find_element_by_id('fal_name').send_keys(father_last_name)
        
        driver.find_element_by_id('faf_name').click()
        
        # get first name
        father_first_name = get_first_name(father_name)
        driver.find_element_by_id('faf_name').send_keys(father_first_name)




        birth_date = str(all_details['Date of Birth'])



        birth_date = birth_date.split('-')
        birth_date = f"{birth_date[2]}/{birth_date[1]}/{birth_date[0]}"

        driver.find_element_by_id('dob').click()
        driver.find_element_by_id('dob').send_keys(birth_date)

        driver.find_element_by_id('tel_num_isdcode').click()
        
        try:
            alert = driver.switch_to.alert
            alert.accept()
            time.sleep(1)
        except:
            pass

        driver.find_element_by_id('tel_num_isdcode').send_keys("91")

        driver.find_element_by_id('tel_num_stdcode').click()
        driver.find_element_by_id('tel_num_stdcode').send_keys("")
        
        driver.find_element_by_id('tel_num').click()
        driver.find_element_by_id('tel_num').send_keys(all_details['Contact No'])


        driver.find_element_by_id('email_id').click()
        driver.find_element_by_id('email_id').send_keys(all_details['Email Id'])

        driver.find_element_by_id('add_comm').click()
        driver.find_element_by_id('add_comm').send_keys('INDIAN')

        driver.find_element_by_id('ra_add').click()

        try:
            time.sleep(1)
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass

        driver.find_element_by_id('ra_add').click()

        driver.find_element_by_id('ra_add').send_keys('INDIAN')
        

        driver.find_element_by_id('proof_id').click()
        sel = Select(driver.find_element_by_id('proof_id'))
        sel.select_by_visible_text("AADHAAR Card issued by the Unique Identification Authority of India")

        driver.find_element_by_id('proof_add').click()
        sel = Select(driver.find_element_by_id('proof_add'))
        sel.select_by_visible_text("AADHAAR Card issued by the Unique Identification Authority of India")
        
        driver.find_element_by_id('proof_dob').click()
        sel = Select(driver.find_element_by_id('proof_dob'))
        sel.select_by_visible_text("AADHAAR Card issued by the Unique Identification Authority of India")

        assessing_office = all_details['Assessing Office']
        assessing_office = assessing_office.replace(" ", "")
        assessing_office = assessing_office.split("|")
        
        if assessing_office[0] == "" or assessing_office[1] == "" or assessing_office[2] == "" or assessing_office[3] == "":
            verify_status = "not done"
            return_message = "nsdl Id me  error - Area Code is mandatory (Apna Pan india me Ao Code Hi  mandatory kra deta hu )"
            return_data =  {"verify_status": verify_status, "message": return_message}

        driver.find_element_by_id('area_code').send_keys(assessing_office[0])
        driver.find_element_by_id('ao_type').send_keys(assessing_office[1])
        driver.find_element_by_id('range_code').send_keys(assessing_office[2])
        driver.find_element_by_id('ao_num').send_keys(assessing_office[3])

        state = all_details['Pan Card Dispatched State']
        print(state)

        driver.find_element_by_id('aoSelection').click()
        
        sel = Select(driver.find_element_by_id('state_aoCode'))
        sel.select_by_visible_text(state)

        sel = Select(driver.find_element_by_id('user_state'))
        sel.select_by_visible_text(str(state))

        state_value = driver.find_element_by_id('user_state').get_attribute('value')
        if state_value == "Please Select":
            time.sleep(1)
            # driver.find_element_by_id('user_state').click()
            # time.sleep(1)
            sel = Select(driver.find_element_by_id('user_state'))
            sel.select_by_visible_text(str(state))


        sel = Select(driver.find_element_by_id('check_aadhaar_eid'))
        sel.select_by_value('A')

        driver.find_element_by_id('aadhaarNo').send_keys(all_details['Aadhar No'])

        driver.find_element_by_id('name_aadhaar').send_keys(all_details['Name on Card'])

        sel = Select(driver.find_element_by_id('gender'))
        sel.select_by_value(all_details['Gender'])

        check_box = driver.find_element_by_id('aadhaarconsent')
        check_box.click()

        driver.find_element_by_id('verify').click()
        time.sleep(2.5)
        try:
            time.sleep(1)
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass
        time.sleep(2.5)
        
        while True:
            return_message = driver.find_element_by_id('result_message').text
            if return_message == 'Demographic verification successful':
                verify_status = "done"

                try:
                    time.sleep(1)
                    alert = driver.switch_to.alert
                    alert.accept()
                except:
                    pass

                driver.find_element_by_id('proceed').click()
                try:
                    time.sleep(1)
                    alert = driver.switch_to.alert
                    alert.accept()
                except:
                    pass

                driver.find_element_by_id('submit').click()
                time.sleep(5)
                # driver.find_element_by_class_name('form_button')[0].click()
                # document.getElementsByClassName('form_button')[0].click()

                break

            elif return_message == 'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.':
                verify_status = "not done"
                break
            
            elif return_message == 'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.':
                verify_status = "not done"
                break

            elif return_message == 'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in':
                verify_status = "not done"
                break
            
            elif return_message == 'Alert Text: Area Code is mandatory\nMessage: unexpected alert open: {Alert text : Area Code is mandatory}\n  (Session info: chrome=121.0.6167.161)\n':
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

    return_data =  {"verify_status": verify_status, "message": return_message}
    print('\n\n')
    print(return_data)
    print('\n\n')

    return return_data

def CorrectionPersonNSDL(all_details):
    try:

        driver.switch_to.window(driver.window_handles[1])
        try:
            driver.find_element_by_id('oCMenu_top3').click()
            time.sleep(0.25)
        except:
            pass
        
        try:
            driver.find_element_by_id('oCMenu_sub12').click()
            time.sleep(0.25)
        except:
            return_data =  {"verify_status": 'not done', "message": 'SESSION EXPIRE'}
            return return_data


        try:
            js = ShowNameAndTokenNo(token_no=all_details['PanCard Serial ID'], Name=all_details['Name'])
            driver.execute_script(js)
        except:
            pass

        driver.find_element_by_id('pan').click()
        driver.find_element_by_id('pan').send_keys = all_details['Old PanCard No']

        pan_number = driver.find_element_by_id('pan').get_attribute('value')

        if pan_number == "":
            js = f"document.getElementById('pan').value = '{all_details['Old PanCard No']}'; console.log('{all_details['Old PanCard No']}')"
            driver.execute_script(js)

        driver.find_element_by_id('category').click()
        sel = Select(driver.find_element_by_id('category'))
        sel.select_by_value(all_details['Category of Applicant'])

        name_list = all_details['Name'].split(" ")

        l_name = name_list[len(name_list) - 1]
        f_name = all_details['Name'].replace(l_name, "")

        driver.find_element_by_id('lastName').click()
        driver.find_element_by_id('lastName').send_keys(l_name)

        driver.find_element_by_id('firstName').click()
        driver.find_element_by_id('firstName').send_keys(f_name)

        father_name = all_details["Father's Last Name"]

        # get last name
        father_last_name = get_last_name(father_name)

        driver.find_element_by_id('fatherlastName').click()
        driver.find_element_by_id('fatherlastName').send_keys(father_last_name)

        driver.find_element_by_id('fatherfirstName').click()

        # get first name
        father_first_name = get_first_name(father_name)

        driver.find_element_by_id('fatherfirstName').send_keys(father_first_name)

        driver.find_element_by_id('crSubmitBtn').click()

        try:
            time.sleep(1)
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass


        driver.find_element_by_id('acceptSubmit').click()

        driver.find_element_by_id('nameOnCard').click()
        driver.find_element_by_id('nameOnCard').send_keys(all_details['Name on Card'])

        birth_date = str(all_details['Date of Birth'])

        birth_date = birth_date.split('-')
        birth_date = f"{birth_date[2]}/{birth_date[1]}/{birth_date[0]}"

        driver.find_element_by_id('dob').click()
        driver.find_element_by_id('dob').send_keys(birth_date)

        driver.find_element_by_id('isd').click()
        driver.find_element_by_id('isd').send_keys("91")

        driver.find_element_by_id('telNum').click()
        driver.find_element_by_id('telNum').send_keys(all_details["Contact No"])

        driver.find_element_by_id('add_comm').click()
        sel = Select(driver.find_element_by_id('add_comm'))
        sel.select_by_value('indian')

        driver.find_element_by_id('emailId').click()
        driver.find_element_by_id('emailId').send_keys(all_details['Email Id'])

        state = all_details['Pan Card Dispatched State']

        driver.find_element_by_id('user_state').click()
        
        sel = Select(driver.find_element_by_id('user_state'))

        sel.select_by_visible_text(str(state))

        driver.find_element_by_id('check_aadhaar_eid').click()
        sel = Select(driver.find_element_by_id('check_aadhaar_eid'))
        sel.select_by_value('A')

        driver.find_element_by_id('aadhaarNo').send_keys(all_details['Aadhar No'])

        driver.find_element_by_id('name_aadhaar').send_keys(all_details['Name on Card'])

        driver.find_element_by_id('poid').click()
        sel = Select(driver.find_element_by_id('poid'))
        sel.select_by_visible_text("AADHAAR Card issued by the Unique Identification Authority of India")

        driver.find_element_by_id('por').click()
        sel = Select(driver.find_element_by_id('por'))
        sel.select_by_visible_text("AADHAAR Card issued by the Unique Identification Authority of India")
        
        driver.find_element_by_id('pod').click()
        sel = Select(driver.find_element_by_id('pod'))
        sel.select_by_visible_text("AADHAAR Card issued by the Unique Identification Authority of India")

        driver.find_element_by_id('gender')
        sel = Select(driver.find_element_by_id('gender'))
        sel.select_by_value(all_details['Gender'])

        check_box = driver.find_element_by_id('aadhaarconsent')
        check_box.click()

        driver.find_element_by_id('verify').click()

        time.sleep(2.5)
        while True:
            return_message = driver.find_element_by_id('result_message').text
            if return_message == 'Demographic verification successful.':
                verify_status = "done"
                driver.find_element_by_id('Save').click()
                time.sleep(0.25)
                driver.find_element_by_id('SaveConfirm').click()
                time.sleep(2.5)
                break

            elif return_message == 'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.':
                verify_status = "not done"
                break
            
            elif return_message == 'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.':
                verify_status = "not done"
                break

            elif return_message == 'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in':
                verify_status = "not done"
                break
            
            elif return_message == 'Alert Text: Area Code is mandatory\nMessage: unexpected alert open: {Alert text : Area Code is mandatory}\n  (Session info: chrome=121.0.6167.161)\n':
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

    return_data =  {"verify_status": verify_status, "message": return_message}
    print('\n\n')
    print(return_data)
    print('\n\n')

    return return_data

def get_latest_file_in_folder(dir_path, file_type):
  
  # Get a list of all the files in the directory with the specified file type
  files = [f for f in os.listdir(dir_path) if f.endswith(file_type)]

  # Create a tuple for each file that includes the file name and its modification time
  file_info = [(f, os.stat(os.path.join(dir_path, f)).st_mtime) for f in files]

  # Sort the list of tuples by the modification time
  file_info.sort(key=lambda x: x[1])

  # Get the last tuple in the sorted list, which will contain the file name and modification time of the latest file
  latest_file = file_info[-1][0]
  # print(latest_file)
  # latest_time = file_info[-1][1]

  # Return the file name and modification time
  return latest_file


def CheckForNewFile(old_pdf_file):
    # download_location = os.path.expanduser("~")+"\\Downloads"
    download_location = os.getcwd() + '\\Downloads'
    while True:
        time.sleep(0.5)
        new_pdf_file = get_latest_file_in_folder(download_location, ".pdf")
        if new_pdf_file == old_pdf_file:
            pass
        else:
            break
    
    return f"{download_location}\\{new_pdf_file}"



def GeneratePNumber(file_name):
    pdf_document = fitz.open(file_name)

    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        text = page.get_text()
        # print(f"Page {page_num + 1}:\n{text}\n")

    pdf_document.close()
    
    text_lsit = text.split('Proof of DOB')
    text_lsit = text_lsit[1].split('Date')
    text = text_lsit[0]
    text = text.replace("\n", "")
    text = text.replace(" ", "")
    text = text.replace("P-", "")
    
    new_name = os.getcwd() + "\\Downloads\\" + text + "hhhh.pdf"
    os.rename(file_name, new_name)
    return text


def EnterNSDLResult(status, message):
    if status == False:
        if message == 'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.':
            # message = 'Aadhar Demographic Error (Name,DOB,Gender,Aadhar No)   "Aadhar data not match"  (ONLINE FORM JO AAP FILL KIYE HAI USME AADHAR NUMBER, NAME,SPELLING,GENDER,SPACE ) YA NEW AADHAR CARD DOWNLOAD KAR CHECK KAREN-'
            message = 'Please Wait Sometime for acknowledgement! Slip'
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])
            sel = Select(driver.find_element_by_id('Status'))
            sel.select_by_value("5")
            
            driver.find_element_by_name('Remarks').clear()
            driver.find_element_by_name('Remarks').send_keys(message)

            driver.find_elements_by_name('submit')[0].click()

        elif message == 'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.':
            message = 'Pan has already issued Been Aadhar (Pan Pahle se Ban gya hai is Aadhar Se)Find To Pan Box No - 04 se Search karen '

            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])
            sel = Select(driver.find_element_by_id('Status'))
            sel.select_by_value("5")
            
            driver.find_element_by_name('Remarks').clear()
            driver.find_element_by_name('Remarks').send_keys(message)

            driver.find_elements_by_name('submit')[0].click()

        elif message == 'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in':
            message = 'Online Aap Aadhar Number Wrong Fill kiye hai... (Check - Online Aadhar Number/Name/DoB/Sppeling)..Wait For Ack Slip'

            time.sleep(1)
            
            driver.switch_to.window(driver.window_handles[0])
            sel = Select(driver.find_element_by_id('Status'))
            # go to processing
            sel.select_by_value("5")
            
            driver.find_element_by_name('Remarks').clear()
            message = message + " Please Wait Sometime for acknowledgement! Slip"
            driver.find_element_by_name('Remarks').send_keys(message)

            driver.find_elements_by_name('submit')[0].click()

        elif 'Alert Text: Area Code is mandatory\nMessage: unexpected alert open: {Alert text : Area Code is mandatory}' in message:
            message = 'Next Time Online AO Code Fill Kiya Karen,Wait karen Ack Slip ke liye.'

            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])
            sel = Select(driver.find_element_by_id('Status'))
            sel.select_by_value("5")
            
            driver.find_element_by_name('Remarks').clear()
            driver.find_element_by_name('Remarks').send_keys(message)

            driver.find_elements_by_name('submit')[0].click()
        

        else:
            time.sleep(1)
            
            driver.switch_to.window(driver.window_handles[0])
            sel = Select(driver.find_element_by_id('Status'))
            sel.select_by_value("5")
            
            driver.find_element_by_name('Remarks').clear()
            message = "Please Wait Sometime for acknowledgement! Slip"
            driver.find_element_by_name('Remarks').send_keys(message)

            driver.find_elements_by_name('submit')[0].click()
    
    if status == True:

        # download_location = os.path.expanduser("~")+"\\Downloads"
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
        except:
            pass

        try:
            time.sleep(2)
            file_name = CheckForNewFile(old_pdf_file)
            p_number = GeneratePNumber(file_name)
            print(p_number)

            driver.find_elements_by_tag_name('input')[1].click()
            driver.switch_to.window(driver.window_handles[0])
            sel = Select(driver.find_element_by_id('Status'))

            sel.select_by_value("1")
            time.sleep(0.5)
            file_name =  os.getcwd() + '\\Downloads\\' + p_number + "hhhh.pdf"
            driver.find_element_by_name('ACkPdf').send_keys(file_name)
            driver.find_element_by_name('ACKNo').send_keys(p_number)
            driver.find_element_by_name('Remarks').clear()
            driver.find_element_by_name('Remarks').send_keys(".")
            driver.find_elements_by_name('submit')[0].click()

            time.sleep(0.25)
            # download the last pdf
            js = "document.getElementsByClassName('btn-final-pdf')[0].click()"
            # driver.find_elements_by_class_name('btn-final-pdf')[0].click()
            driver.execute_script(js)

        except:
            driver.switch_to.window(driver.window_handles[1])
            command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""Please Fill This PDF Manually, and After Filling Click on Home Button"", 0, ""Information"":close")'
            subprocess.Popen(command)
            time.sleep(2)
            
            while True:
                time.sleep(1)
                url = driver.current_url
                if "onlineservices.nsdl.com/paam/homeTinFC.html?ID" in url:
                    print('i start my work again')
                    time.sleep(3)

                    driver.switch_to.window(driver.window_handles[0])
                    break
                
                else:
                    continue



InitiliseBrowser()
apna_pan_login_status = LoginApnaPan()

g = LoginNSDL()


global removed_states
removed_states = ['ASSAM', 'MIZORAM', 'JAMMU AND KASHMIR']

if g == True:

    while True:
        try:
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])
        except:
            current_url = driver.current_url
            if "apnapanindia.co.in" in current_url:
                pass
            else:
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[0])
        
        driver.get('https://apnapanindia.co.in/soft/SuperAdmin/pancard-power/op-index.php')
        sel = Select(driver.find_elements_by_name('PanType')[0])
        sel.select_by_value('Individual')
        sel = Select(driver.find_elements_by_name('Status')[0])
        sel.select_by_value('4')
        driver.find_elements_by_class_name('filtr')[0].click()
        try:
            setPage = driver.find_elements_by_class_name('setPage')[0].text
            setPage = setPage.replace(" ", "")
            setPage = setPage.split("of")
            setPage = setPage[1]
            
            url = f'https://apnapanindia.co.in/soft/SuperAdmin/pancard-power/op-index.php?page={setPage}'
            driver.get(url)
        except:
            pass
    
        btns = driver.find_elements_by_class_name('DetailBTN')
        
        all_urls = []
        
        for btn in btns:
            url = btn.get_attribute('href')
            all_urls.append(url)
        
        if all_urls == []:
            command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""All work is completed..."", 0, ""Information"":close")'
            driver.window_handles[1]
            current_url = driver.current_url
            driver.get(current_url)
            subprocess.Popen(command)
            driver.quit()
            exit()

            

        for url in all_urls:

            driver.get(url)
            html_content = driver.find_elements_by_tag_name('table')[0].get_attribute('outerHTML')
            
            # html_content = driver.page_source
            df = pd.read_html(html_content)[0] 
            df = df.to_dict(orient='records')
            page_data = {}
            for data in df:
                page_data[data[0]] = str(data[1])

            if page_data['Pan Card Dispatched State'] in removed_states:
                message = "Please wait for Acknowledgement slip in PAN card."
                work_status = {}
                work_status['verify_status'] = 'not done'
                work_status['message'] = message
                

            elif page_data['PanCard Type'] == 'New':
                
                for n in range(3):
                    work_status = NewPersonNSDL(page_data)
                    print(work_status)
                    
                    if work_status['verify_status'] == "not done":
                        print(work_status['message'])

                        if work_status['message'] == 'SESSION EXPIRE':
                            command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""Your Session is Expired..."", 0, ""Information"":close")'
                            subprocess.Popen(command)
                            driver.quit()
                    
                        elif work_status['message'] == 'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.':
                            break
                        
                        elif work_status['message'] == 'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.':
                            break

                        elif work_status['message'] == 'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in':
                            break
                        
                        elif work_status['message'] == 'Alert Text: Area Code is mandatory\nMessage: unexpected alert open: {Alert text : Area Code is mandatory}\n  (Session info: chrome=121.0.6167.161)\n':
                            break
                    

                        try:
                            alert = driver.switch_to.alert
                            alert.accept()
                        except:
                            pass
                    else:
                        break
            

            elif page_data['PanCard Type'] == 'Correction':
                
                for n in range(3):
                    work_status = CorrectionPersonNSDL(page_data)

                    if work_status['verify_status'] == "not done":
                        if work_status['message'] == 'SESSION EXPIRE':
                            command=f'mshta vbscript:Execute("CreateObject(""WScript.Shell"").Popup ""Your Session is Expired..."", 0, ""Information"":close")'
                            driver.window_handles[1]
                            current_url = driver.current_url
                            driver.get(current_url)
                            subprocess.Popen(command)
                            driver.quit()
                            exit()

                        elif work_status['message'] == 'Demographic authentication failed as the details (Name, DOB & Gender) entered by you are not matching with the details available in UIDAI database.Please recheck the details entered by you. If there is any error, then please capture details once again to generate a new receipt;If the details entered are correct, please proceed with biometric authentication using biometric device installed at your Centre (Protean TIN-FC/PAN Centre) by selecting biometric option shown above.':
                            break
                        
                        elif work_status['message'] == 'On verification with PAN database, it appears that PAN has already been issued against Aadhaar quoted in the application form. Please advise applicant to submit his/her application using PAN Change Request form quoting his/her PAN.':
                            break

                        elif work_status['message'] == 'Demographic verification error,Please try after some time.If error persists,then kindly contact paam@nsdl.co.in':
                            break
                        
                        elif work_status['message'] == 'Alert Text: Area Code is mandatory\nMessage: unexpected alert open: {Alert text : Area Code is mandatory}\n  (Session info: chrome=121.0.6167.161)\n':
                            break
                        
                        try:
                            alert = driver.switch_to.alert
                            alert.accept()
                        except:
                            pass
                    else:
                        break
            

            if work_status['verify_status'] == "done":
                EnterNSDLResult(True, ".")

            
            if work_status['verify_status'] == 'not done':
                EnterNSDLResult(False, work_status['message'])

            


            




    

# driver.findElementsByTagName('a')[0].text == 'Tax Information Network'

