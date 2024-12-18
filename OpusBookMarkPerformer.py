import json
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import os
from openpyxl import load_workbook
from OpenOrchestrator.database.queues import QueueElement
from datetime import datetime
import calendar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import os
import time
import requests

#Connect to orchestrator
orchestrator_connection = OrchestratorConnection("Opus bookmark performer", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)

log = False

if log:
    orchestrator_connection.log_info("Started process")

#Opus bruger
OpusLogin = orchestrator_connection.get_credential("OpusBruger")
OpusUser = OpusLogin.username
OpusPassword = OpusLogin.password 

#Robotpassword
RobotCredential = orchestrator_connection.get_credential("Robot365User") #### opusbruger??
RobotUsername = RobotCredential.username
RobotPassword = RobotCredential.password

# Define the queue name
queue_name = "OpusBookmarkQueue"  # Replace with your queue name

 # Assign variables from SpecificContent
PasswordString = None
OpusBookmark = None
SharePointURL = None
FileName = None
Daily = None
MonthEnd = None
MonthStart = None
Yearly = None
i = 1
# Get all queue elements with status 'New'
queue_item = orchestrator_connection.get_next_queue_element(queue_name)
if not queue_item:
    orchestrator_connection.log_info("No new queue items to process.")
    exit()

specific_content = json.loads(queue_item.data)

if log:
    orchestrator_connection.log_info("Assigning variables")
# Assign variables from SpecificContent
PasswordString = OpusPassword      ############# Måske ikke rigtigt
BookmarkID = specific_content.get("Bookmark")
OpusBookmark = orchestrator_connection.get_constant("OpusBookMarkUrl").value + BookmarkID
SharePointURL = orchestrator_connection.get_constant("LauraTestSharepointURL").value + "/Delte Dokumenter/"
#SharepointURL = specific_content.get("SharePointMappeLink", None)
FileName = specific_content.get("Filnavn", None)
Daily = specific_content.get("Dagligt (Ja/Nej)", None)
MonthEnd = specific_content.get("MånedsSlut (Ja/Nej)", None)
MonthStart = specific_content.get("MånedsStart (Ja/Nej)", None)
Yearly = specific_content.get("Årligt (Ja/Nej)", None)
print(BookmarkID, OpusBookmark, FileName, Daily, MonthEnd, MonthStart, Yearly)
# Mark the queue item as 'In Progress'
orchestrator_connection.set_queue_element_status(queue_item.id, "IN_PROGRESS")

# Mark the queue item as 'Done' after processing
orchestrator_connection.set_queue_element_status(queue_item.id, "DONE")

Run = False

#Testing if it should run
if Daily == "ja":
    Run = True
else:
    current_date = datetime.now()
    year, month, day = current_date.year, current_date.month, current_date.day
    
    # Check for month-end
    last_day_of_month = calendar.monthrange(year, month)[1]  
    if MonthEnd == "ja" and day == last_day_of_month:
        Run = True
    # Check for month-start
    elif MonthStart == "ja" and day == 1:
        Run = True
    # Check for year-end
    elif Yearly == "ja" and day == 31 and month == 12:
        Run = True
    
if Run:
    print("Running")
    # Delete the file if it exists
    if os.path.exists(FileName):
        os.remove(FileName)

    # SharePoint credentials
    if log:
        orchestrator_connection.log_info("Connecting to sharepoint")
    SharePointcreds = orchestrator_connection.get_credential("GraphAppIDAndTenant")
    SharePointAppID = SharePointcreds.username
    SharePointTenant = SharePointcreds.password
    SharepointURL_connection = orchestrator_connection.get_constant("AktbobSharePointURL").value
    SharepointURL_connection = orchestrator_connection.get_constant("LauraTestSharepointURL").value

    #Connecting to sharepoint
    credentials = UserCredential(RobotUsername, RobotPassword)
    ctx = ClientContext(SharepointURL_connection).with_credentials(credentials)
    #Checking connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    # Selenium configuration
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
    })
    chrome_service = Service()  # Dynamically locate ChromeDriver if required
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    try:
        # Step 1: Navigate to the Opus portal and log in
        driver.get(orchestrator_connection.get_constant("OpusAdgangUrl").value)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
        
        username_field = driver.find_element(By.ID, "logonuidfield")
        password_field = driver.find_element(By.ID, "logonpassfield")
        username_field.send_keys(OpusUser)
        password_field.send_keys(OpusPassword)
        driver.find_element(By.ID, "buttonLogon").click()

        # Step 2: Navigate to the specific bookmark
        driver.get(OpusBookmark)
        WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[id^='iframe_Roundtrip']"))
        )

        # Step 3: Wait for the export button to appear
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "ACTUAL_DATE_TEXT_TextItem"))
        )
        driver.find_element(By.ID, "BUTTON_EXPORT_btn1_acButton").click()

        # Step 4: Wait for the file download to complete
        initial_file_count = len(os.listdir(downloads_folder))
        start_time = time.time()
        while True:
            files = os.listdir(downloads_folder)
            if len(files) > initial_file_count:
                latest_file = max(
                    [os.path.join(downloads_folder, f) for f in files], key=os.path.getctime
                )
                if latest_file.endswith(".xls"):
                    break
            if time.time() - start_time > 1800:  # Timeout after 30 minutes
                raise TimeoutError("File download did not complete within 30 minutes.")
            time.sleep(1)

        # Step 5: Convert the downloaded file to .xlsx
        xls_file_path = latest_file
        xlsx_file_path = os.path.join(downloads_folder, FileName + ".xlsx")

        wb = load_workbook(xls_file_path)
        wb.save(xlsx_file_path)
        os.remove(xls_file_path)

    except Exception as e:
        orchestrator_connection.log_error(f"An error occurred during Selenium operations: {str(e)}")
    finally:
        driver.quit()
    
    if log:
        orchestrator_connection.log_info("Getting file/folder")
    

    file_name = xlsx_file_path.split("/")[-1]
    download_path = os.path.join(os.getcwd(), file_name)

    # Download the file
    with open(download_path, "wb") as local_file:
        file = ctx.web.get_file_by_server_relative_path(xlsx_file_path).download(local_file).execute_query()

    if log:
        orchestrator_connection.log_info("Uploading file to sharepoint")

    # Send the GET request to SharePoint
    response = requests.get(SharePointURL)
    response.raise_for_status()

    # Parse the JSON response
    folder_info = response.json()

    target_url = f"{SharePointURL}/{FileName}"

    # The upload
    with open(xlsx_file_path, "rb") as local_file:
        target_folder = ctx.web.get_folder_by_server_relative_url(SharePointURL)
        target_folder.upload_file(file_name, local_file.read()).execute_query()

    #Removing the local file
    if os.path.exists(FileName):
        os.remove(FileName)

        

