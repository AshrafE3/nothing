    # Libraries
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import getpass
import time
import os
try:
    from pygame import mixer
except:
    os.system("pip install pygame")
finally:
    from pygame import mixer
try:
    import docx
except:
    os.system("pip install python-docx")
finally:
    import docx

# Function to check if the filename contains one digit or one letter
def containsDigitOrLetter(strr, current_dir):
    name_extension = strr.split("\\")[-1]
    onlyname = strr.split("\\")[-1].split(".")[0]
    for char in onlyname:
        if char.isdigit() or char.isalpha():
            return True
    errors_folder = os.path.join(current_dir, "Errors")
    if not os.path.exists(errors_folder):
        os.mkdir(errors_folder)
    os.rename(strr, os.path.join(errors_folder, name_extension))
    return False

# Upload files
def uploadFiles(files, mp3_file):
    driver.get("https://www.slideshare.net/upload")
    time.sleep(.1)
    upload_files = wait.until(EC.presence_of_element_located((By.ID, 'local-upload'))).send_keys("\n".join(files))
    time.sleep(.2)
    # Tries To solve Captcha
    if "Verify that you are human" in driver.find_element(By.TAG_NAME, "body").text:
        print("\nTrying To Solve Captcha")
        iframe = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'iframe[title="reCAPTCHA"]')))
        driver.switch_to.frame(iframe)
        captcha = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.recaptcha-checkbox-border")))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable(captcha))
        captcha.click()
        driver.switch_to.default_content()
    time.sleep(.2)

    # Solve It Manually if it's not solved
    while "Verify that you are human" in driver.find_element(By.TAG_NAME, "body").text:
        if mp3_file:
            mixer.music.play()
        input("\nFailed To Solve Captcha, Please Solve It And Then Press Enter After Finishing..")
        time.sleep(.1)
    else:
        print("\nCaptcha Solved!")
    time.sleep(.2)
    if "Something went wrong, try again" in driver.find_element(By.TAG_NAME, "body").text:
        print("\n\nUnexpected Error! Maybe You're Not Logged in! Exiting...")
        if mp3_file:
            mixer.music.play()
        time.sleep(.3)
        driver.quit()
        exit(1)
    # Wait Until Files Are Uploaded
    
    print('waiting file upload')
    while len(files) != len(driver.find_elements(By.CSS_SELECTOR, 'div[class*="UploadForm_progressDetailsContainer"] [class*="UploadForm_uploadProgress"]')):
        time.sleep(1)
    time.sleep(3)
    while True:
        try:
            driver.find_element(By.CSS_SELECTOR, '[data-testid="progress-stroke"]')
            time.sleep(1)
        except:
            break
    print('complete file upload')
    
    #h6_elements = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "h6")))
    #for h6 in h6_elements:
    #    while " Mb" not in h6.text:
    #        time.sleep(0.2)

def fillContent(files):
    for x in range(len(files)):
        categories = driver.find_elements(By.ID, 'category')
        categories[x].send_keys("Education")
        descriptions = driver.find_elements(By.ID, 'description')
        description = descriptions[x]
        filetext = getText(files[x])
        driver.execute_script("arguments[0].value+=arguments[1]", description, filetext)
        time.sleep(.2)
        description.send_keys(".")
        time.sleep(.1)

def getText(filename):
    doc = docx.Document(filename)
    fullText_list = []
    for para in doc.paragraphs:
        fullText_list.append(para.text.strip())
    fullText = '\n'.join(fullText_list)
    try:
        fullText = fullText[:2996]
    except:
        pass
    return fullText

def publishFiles(counter):
    publish_buttons = driver.find_elements(By.XPATH, '//button[@class="undefined Button_button__j46GA Button_primary__nZjNo "]')
    for button in publish_buttons:
        print(counter)
        button.click()
        time.sleep(1)
        counter+=1
    for button in publish_buttons:
        wait.until(EC.staleness_of(button))
    return counter

def deleteFiles(files):
    for file in files:
        os.remove(file)

#def editTitles():
 #   titles = driver.find_elements(By.ID, "title")
 #   for title in titles:
 #       current_text = title.get_attribute("value")
 #       title.clear()
 #       title.send_keys(current_text.replace(".docx", ""))

# get current working directory
current_dir = os.getcwd()
if os.path.exists("Input"):
    current_dir = os.path.join(current_dir, "Input")
else:
    print("Directory Name 'Input' is Not Found, Exiting...")
    exit(1)

if os.path.exists("beep-05.mp3"):
    mp3_file = os.path.join(os.getcwd(), "beep-05.mp3")
    mixer.init()
    mixer.music.load(mp3_file)
else:
    mp3_file = ""
    print("\nbeep-05.mp3 File Not Found, Notifications Will Be Disabled.." )
    x = ""
    while x.lower() != "y" and x.lower() != "n":
        x = input("\nAre You Sure You Want To Continue Without Notifications?[y/n]: ")
    if x == "n":
        while not os.path.exists(mp3_file) and ".mp3" not in mp3_file:
            mp3_file = os.path.join(os.getcwd(), input("Please Enter The New Beep Sound File (Absolute Path): "))
        mixer.init()
        mixer.music.load(mp3_file)

# get subfolders
folders = [os.path.join(current_dir, name) for name in os.listdir(current_dir) if os.path.isdir(os.path.join(current_dir, name))]
if not folders:
    print(f"\nNo Directories Found In The Path: {current_dir}")
    exit(1)

# Get login credentials
email = input("Please Enter Email Address: ")
password = input("Please Enter Password: ")

# add options to disable password pop up
options = webdriver.ChromeOptions()
options.add_argument("start-maximized")
prefs = {"credentials_enable_service": False, "profile.password_manager_enabled": False}
options.add_experimental_option("prefs", prefs)

# launch driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)


# open the url

url = "https://www.slideshare.net/"
driver.get(url)

# login
wait.until(EC.presence_of_element_located((By.ID, 'login-from-header'))).click()
wait.until(EC.presence_of_element_located((By.ID, 'scribd-login'))).click()

sign_in_1 = wait.until(EC.presence_of_element_located((By.XPATH, '//a[@class="wrapper__text_button last_lb_item"]')))
driver.execute_script("arguments[0].click();", sign_in_1)

sign_in_2 = wait.until(EC.presence_of_element_located((By.XPATH, '//a[@class="wrapper__text_button _2IxBKS"]')))
driver.execute_script("arguments[0].click();", sign_in_2)

email_input = wait.until(EC.presence_of_element_located((By.ID, "login_or_email")))
wait.until(EC.element_to_be_clickable(email_input))
email_input.send_keys(email)
time.sleep(1)
password_input = wait.until(EC.presence_of_element_located((By.ID, "login_password")))
wait.until(EC.element_to_be_clickable(email_input))
password_input.send_keys(password)
time.sleep(2)
login_button = wait.until(EC.presence_of_element_located((By.XPATH, '//button[@class="wrapper__filled-button _5fglnc"]')))
wait.until(EC.element_to_be_clickable(login_button))
driver.execute_script("arguments[0].click();", login_button)
time.sleep(3)
counter=0
while "Share what you know and love through presentations, infographics, documents and more" not in driver.find_elements(By.TAG_NAME, 'body')[0].text:
    if "Welcome back!" in driver.find_elements(By.TAG_NAME, 'body')[0].text or "Logging you in..." in driver.find_elements(By.TAG_NAME, 'body')[0].text:
        break
    time.sleep(0.2)
    counter+=0.2
    if counter>5:
        if "Welcome back!" in driver.find_elements(By.TAG_NAME, 'body')[0].text or "Logging you in..." in driver.find_elements(By.TAG_NAME, 'body')[0].text:
            break
        if mp3_file:
            mixer.music.play()
        input("Login Button Clicked, Solve Captcha If Exists, Press Enter After Finishing..")
time.sleep(3)
input("*******Press Enter To Continue******")
# fill the title, description and publish.
counter = 1
for folder in folders:
    print("Folder:", folder.split("\\")[-1])
    files = [os.path.join(folder, name) for name in os.listdir(folder) if os.path.isfile(os.path.join(folder, name)) and containsDigitOrLetter(os.path.join(folder, name), current_dir)]
    #for x in files:
    #    print(x)
    if files:
        uploadFiles(files, mp3_file)
        time.sleep(2)
        #editTitles()
        fillContent(files)
        time.sleep(10)
        counter = publishFiles(counter)
        deleteFiles(files)
        time.sleep(2)
        print("\n\n All Are Done !! \n\n")
        os.rmdir(folder)
        input("*******Change the Account and restart the program******")
    else:
        print(f"\nNo DOCX Files Found In The Directory: {folder}")

driver.quit()

print("Done Processing All Files..")

if mp3_file:
    mixer.music.play()