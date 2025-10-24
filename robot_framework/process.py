"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import pandas as pd
import requests

import time
import os


# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    

    sharepoint_site = f"{orchestrator_connection.get_constant("AarhusKommuneSharePoint").value}/Teams/tea-teamsite10946"
    # Setup Selenium WebDriver

    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")
    
    tenant = api.username
    client_id = api.password
    thumbprint = certification.username
    cert_path = certification.password
    
    client = sharepoint_client(tenant, client_id, thumbprint, cert_path, sharepoint_site, orchestrator_connection)


    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder
    })
    options.add_argument('--remote-debugging-pipe')
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 60)

    # Navigate to the URL
    token = orchestrator_connection.get_credential("VejmanToken").password
    driver.get(f"https://vejman.vd.dk/query/default.do?Parent=6&Item=683&token={token}")

    # Wait for page to load
    wait.until(EC.text_to_be_present_in_element((By.ID, "lastUpdated"), "Senest opdateret"))

    # Click Advanced Search Button
    driver.find_element(By.ID, "tabbtnAdvQuery").click()

    # Enter SQL Query
    query_textbox = driver.find_element(By.CLASS_NAME, "AdvancedSearchVqlTextArea")
    current_year = time.strftime("%Y")
    sql_query = f"""
    VÆLG
    TILL_Sags_Id

    HVOR
    TILL_Bestyrer = 'Aarhus' OG
    TILL_År >= 2024 OG
    TILL_År <= {current_year} OG
    (TILL_Type = 'Grave' ELLER
    TILL_Type = 'Materiel')

    UDVID
    (*
    SELECT DISTINCT ca.pm_case_id "SagsID",
    (select typ.text from h_pm_type typ where typ.id =ca.pm_type) "Type",
    (select stat.text from h_pm_state stat where stat.id =ca.pm_state) "Status", 
    (select (pm_concat(pm_site.street_name)) from pm_site where pm_site.pm_case_id = ca.pm_case_id ) "Vejnavn",
    vql_permission.applicant_name(ca.pm_case_id) "Ansøger",
    vql_permission.contractor_name(ca.pm_case_id) "Entreprenør",
    vql_permission.proprietor_name(ca.pm_case_id) "Ledningsejer",
    ca.user_name "Sagsbehandler", 
    to_char(ca.created_date,'dd-mm-yyyy') "Ansøgningsdato", 
    to_char(ca.start_date, 'dd-mm-yyyy') "Startdato", 
    (ca.start_date-ca.created_date) "Dage fra ansøgning til start",
    vql_permission.Approval_Date(ca.pm_case_id) "Godkendtdato",
    vql_permission.Rejection_Date(ca.pm_case_id) "Afvistdato",
    (vql_permission.dApproval_Date(ca.pm_case_id) - ca.created_date) "Dage fra ansøgning til godkendt",
    (ca.start_date - (vql_permission.dApproval_Date(ca.pm_case_id))) "Dage fra godkendt til start",
    (vql_permission.dRejection_Date(ca.pm_case_id) - ca.created_date) "Dage fra ansøgning til afvisning",
    ca.authority_reference_number "Intern bemærkning"
    FROM visql_tabel vt, pm_case ca
    WHERE 
    vt.TILL_SAGS_ID = ca.Pm_case_id    
    ORDER BY "Ansøgningsdato"
    *)"""

    query_textbox.clear()
    driver.execute_script("arguments[0].value = arguments[1];", query_textbox, sql_query)
    time.sleep(5)

    # Click Search Button
    driver.find_element(By.ID, "btnAdvSearch").click()

    # Wait for result
    driver.implicitly_wait(20)
    wait.until(EC.text_to_be_present_in_element((By.ID, "advResultattabel"), "Resultattabellen er for stor"))

    initial_files = set(os.listdir(downloads_folder))

    # Click Export Button
    driver.find_element(By.XPATH, "//input[@type='submit' and @value='Eksportér']").click()

    # Wait for file to be downloaded
    timeout = 360
    start_time = time.time()


    while True:
        # Get the current list of files
        current_files = set(os.listdir(downloads_folder))
        new_files = current_files - initial_files
        
        # Check if new files have been added
        if new_files:
            # Filter for .xlsx files among the new files
            xlsx_files = [file for file in new_files if file.lower().endswith(".xlsx")]
            if xlsx_files:
                downloaded_file = os.path.join(downloads_folder, xlsx_files[0])
                orchestrator_connection.log_info(f"Download completed: {downloaded_file}")
                break
        
        # Check for timeout
        if time.time() - start_time > timeout:
            orchestrator_connection.log_info("Timeout reached while waiting for a download.")
            break
        
        time.sleep(1)  # Avoid hammering the file system

    # Convert to XLSX
    xlsx_filepath = os.path.join(downloads_folder, "VejmanBehandlingstider.xlsx")
    if os.path.exists(xlsx_filepath):
        os.remove(xlsx_filepath)

   # Remove existing version if it exists
    if os.path.exists(xlsx_filepath):
        os.remove(xlsx_filepath)

    # ✅ Just rename instead of reading/saving
    os.rename(downloaded_file, xlsx_filepath)
    orchestrator_connection.log_info(f"Renamed downloaded file to: {xlsx_filepath}")

    # Cleanup browser
    driver.quit()


    sharepoint_folder = "Delte dokumenter/DataudtrækVejman"

    upload_file_to_sharepoint(client, sharepoint_folder, xlsx_filepath, orchestrator_connection)
    os.remove(xlsx_filepath)

    #Get GraveMaterielTilladelser
    URL = f"https://vejman.vd.dk/permissions/getcases?pmCaseStates=1%2C2&pmCaseFields=state%2Ctype%2Ccase_number%2Cauthority_reference_number%2Ccreated_date%2Cwebgtno%2Cstreet_name%2Capplicant%2Cproprietor%2Ccontractor%2Cinitials&pmCaseWorker=all&pmCaseTypes=%27gt%27%2C%27rovm%27&pmCaseVariant=all&pmCaseTags=ignorerTags&pmCaseTagShow=&pmCaseShowAttachments=false&pmAllStates=&dontincludemap=1&cse=&policeDistrictShow=&_={int(time.time()*1000)}&token={token}"
    # Fetch JSON data
    response = requests.get(URL)
    json_data = response.json()

    # Extract cases
    data = json_data.get("cases", [])

    # Convert JSON to a DataFrame
    df = pd.DataFrame(data)

    # Write to Excel
    xlsx_filepath = "GraveMaterielTilladelser.xlsx"
    with pd.ExcelWriter(xlsx_filepath, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cases")
        
    upload_file_to_sharepoint(client, sharepoint_folder, xlsx_filepath, orchestrator_connection)
    os.remove(xlsx_filepath)


def sharepoint_client(tenant: str, client_id: str, thumbprint: str, cert_path: str, sharepoint_site_url: str, orchestrator_connection: OrchestratorConnection) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    cert_credentials = {
        "tenant": tenant,
        "client_id": client_id,
        "thumbprint": thumbprint,
        "cert_path": cert_path
    }
    ctx = ClientContext(sharepoint_site_url).with_client_certificate(**cert_credentials)

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    orchestrator_connection.log_info(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx

def upload_file_to_sharepoint(client: ClientContext, sharepoint_file_url: str, local_file_path: str, orchestrator_connection: OrchestratorConnection):
    """
    Uploads the specified local file back to SharePoint at the given URL.
    Uses the folder path directly to upload files.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = path_parts[1]
    file_name = os.path.basename(local_file_path)  # File name

    # Construct the server-relative folder path (starting with the document library)
    if FOLDER_PATH:
        folder_path = f"{DOCUMENT_LIBRARY}/{FOLDER_PATH}"
    else:
        folder_path = f"{DOCUMENT_LIBRARY}"

    # Get the folder where the file should be uploaded
    target_folder = client.web.get_folder_by_server_relative_url(folder_path)
    client.load(target_folder)
    client.execute_query()

    # Upload the file to the correct folder in SharePoint
    with open(local_file_path, "rb") as file_content:
        uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()


    orchestrator_connection.log_info(f"[Ok] file has been uploaded to: {uploaded_file.serverRelativeUrl} on SharePoint")
