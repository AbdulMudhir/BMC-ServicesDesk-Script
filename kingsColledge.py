from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, \
    StaleElementReferenceException, NoSuchWindowException, TimeoutException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains

options = Options()
options.add_argument("--user-data-dir=chrome-data")

driver = webdriver.Chrome("chromedriver.exe", options=options)

driver.implicitly_wait(10)

driver.get(
    "https://kcl--bmcservicedesk.eu29.visual.force.com/apex/RemedyforceConsole?sfdc.tabName=01r20000000gKRo#false")


def select_ticket(driver):
    # element for the first row
    first_row_css_selector = "#gridview-1048 > table > tbody > tr:nth-child(2) > td.x-grid-cell.x-grid-cell-gridcolumn-1033 > div"

    try:  # selects and clicks the first row in the ticket queue
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, first_row_css_selector))).click()
    except StaleElementReferenceException:
        sleep(2)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, first_row_css_selector))).click()
    except StaleElementReferenceException:
        sleep(5)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, first_row_css_selector))).click()


def update_ticket(driver):
    resolution_message = "no response from user"

    resolution_text_field_id = "thpage:theForm:thePAgeBlock:j_id84:1:pageSectionId:j_id86:12:j_id87:inputField"
    # ticket main frame
    try:
        main_frame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,    "// iframe[contains( @ id, 'a1N3')]")))
    except TimeoutException:
        sleep(10)
        main_frame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "// iframe[contains( @ id, 'a1N3')]")))

    # switch to the main ticket frame
    driver.switch_to.frame(main_frame)
    # within the main frame there's a child frame moved into
    resolution_frame = driver.find_element_by_id("incidentDetailsFrameId")
    driver.switch_to.frame(resolution_frame)

    # scroll half way through the page
    scroll = ActionChains(driver).send_keys(Keys.PAGE_DOWN)

    scroll.perform()

    # find resolution text field
    try:
        resolution_text_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, resolution_text_field_id)))
        # set the info
        resolution_text_field.send_keys(resolution_message)

    except NoSuchWindowException:
        sleep(5)
        scroll.perform()
        resolution_text_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, resolution_text_field_id)))
        # set the info
        resolution_text_field.send_keys(resolution_message)

    try:
        update_ticket_status = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "statusSelectId"))).click()
        resolve_state = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#statusSelectId > option:nth-child(11)"))).click()

    except ElementClickInterceptedException:
        sleep(2)
        update_ticket_status = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "statusSelectId"))).click()
        resolve_state = WebDriverWait(driver, 5).until( EC.presence_of_element_located((By.CSS_SELECTOR, "#statusSelectId > option:nth-child(11)"))).click()


    ActionChains(driver).send_keys(Keys.ESCAPE).perform()

    driver.switch_to.default_content()

    drop_down_menu_frame = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "// iframe[contains( @ id, 'a1N3')]")))
    driver.switch_to.frame(drop_down_menu_frame)

    save_ticket = driver.find_element_by_id("saveId-btnEl").click()

    sleep(5)



def send_email(driver):
    # driver.switch_to.default_content()
    # drop_down_menu_frame = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#iframe-a1N3X00000QOIxa")))
    # driver.switch_to.frame(drop_down_menu_frame)

    try:
        action_to_action = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "actionButton-btnEl"))).click()
    except ElementClickInterceptedException:
        sleep(2)
        action_to_action = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "actionButton-btnEl"))).click()
    except ElementClickInterceptedException:
        sleep(5)
        action_to_action = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "actionButton-btnEl"))).click()

    select_email_drop_down = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "emailId-itemEl"))).click()

    switch_to_email_window = driver.switch_to.window(driver.window_handles[1])

    email_text_field_css = "#j_id0\:mailForm\:mailBodyRich\:textAreaDelegate_RichTextNote__c_rta_body"

    message_frame = driver.find_element_by_css_selector("#cke_1_contents > iframe")

    driver.switch_to.frame(message_frame)

    email_message_field = driver.find_element_by_css_selector(
        "#j_id0\:mailForm\:mailBodyRich\:textAreaDelegate_RichTextNote__c_rta_body")

    email_message_template = '''
    
Hello,

We haven’t heard back from you about any issues you might be having with your new laptop, and are now resolving this ticket.

If you are having any IT issues or have any questions, please do contact the Service Desk on 0207 848 8888 or by going to 88888.kcl.ac.uk where you can raise a request online.

Kind regards
Anik Talukder'''

    email_message_field.send_keys(email_message_template)

    driver.switch_to.default_content()

    send_email_button = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "BTNsendmailId"))).click()


def move_to_next_ticket(driver):

    sleep(2)

    window = driver.window_handles[0]

    driver.switch_to.window(window)

    sleep(2)
    close_current_ticket = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "x-tab-close-btn"))).click()


def start_finding_ticket(driver):

    print("Selecting first ticket in the queue")
    select_ticket(driver)
    print("Updating the queue")
    update_ticket(driver)
    print("Sending user Email")
    send_email(driver)
    print("moving to next Queue")
    move_to_next_ticket(driver)


def check_validation(driver):

    ticket_owner = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ticket_owner_css))).text

    ticket_status = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ticket_status_css))).text

    if ticket_owner == "Anik Talukder" and ticket_status == "CUSTOMER DEFERRED":

       start_finding_ticket(driver)

    elif ticket_owner == "Anik Talukder" and ticket_status != "CUSTOMER DEFERRED":
        # sleep for 5 seconds to wait for filter to be applied
        sleep(2)

        print("This ticket status is not CUSTOMER DEFERRED .... pending update for filter to take place")

        ticket_status = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ticket_status_css))).text

        if ticket_status != "CUSTOMER DEFERRED":
            print("Ticket states is still NOT CUSTOMER DEFERRED something is wrong... attempting to change filter")
            hover_to_filter_bar.move_to_element(
                driver.find_element_by_id("gridcolumn-1045-triggerEl")).click().perform()
            # filter the ticket by waiting for staff
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#BMCServiceDesk__FKStatus__c_cr1DivId > label:nth-child(4) > span"))).click()
            # apply the filter
            driver.find_element_by_css_selector("#BMCServiceDesk__FKStatus__c_apply").click()
            sleep(5)
            ticket_owner = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ticket_owner_css))).text

            ticket_status = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ticket_status_css))).text

            if ticket_owner == "Anik Talukder" and ticket_status != "CUSTOMER DEFERRED":
                print("I am shutting down.. something is wrong. report this")
                driver.close()

        else:

            start_finding_ticket(driver)
    else:
        print("I am shutting down.. Unknown ticket owner, report this!")
        driver.close()


print("Changing ticket filter to CUSTOMER DEFFERRED")
status = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "gridcolumn-1045-titleEl")))
hover_to_filter_bar = ActionChains(driver).move_to_element(status)
hover_to_filter_bar.move_to_element(driver.find_element_by_id("gridcolumn-1045-triggerEl")).click().perform()

ticket_owner_css = "#gridview-1048 > table > tbody > tr:nth-child(2) > td.x-grid-cell.x-grid-cell-gridcolumn-1040 > div"

ticket_status_css = "#gridview-1048 > table > tbody > tr:nth-child(2) > td.x-grid-cell.x-grid-cell-gridcolumn-1045 > div"

try:
# filter the ticket by waiting for staff
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#BMCServiceDesk__FKStatus__c_cr1DivId > label:nth-child(4) > span"))).click()
    # apply the filter
    driver.find_element_by_css_selector("#BMCServiceDesk__FKStatus__c_apply").click()
except ElementNotInteractableException:
    status = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "gridcolumn-1045-titleEl")))
    hover_to_filter_bar = ActionChains(driver).move_to_element(status)
    hover_to_filter_bar.move_to_element(driver.find_element_by_id("gridcolumn-1045-triggerEl")).click().perform()
    sleep(5)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "#BMCServiceDesk__FKStatus__c_cr1DivId > label:nth-child(4) > span"))).click()
    # apply the filter
    driver.find_element_by_css_selector("#BMCServiceDesk__FKStatus__c_apply").click()

while 1:
    check_validation(driver)