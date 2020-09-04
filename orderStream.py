## ---------------------------------------------------------------------------------------------------------- #
""" \package orderStream
    \brief   This script performs automated tasks on the OrderStream sales management website. There are 
             two provided functions, thus far: process/invoice open orders, and extract merchant id's. For 
             either modes, login information is required at the command line. To perform merchant id 
             extraction, give the script the id flag.

"""
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver

import argparse
import enum
import pandas
import random
import sys
import time


## ---------------------------------------------------------------------------------------------------------- #
# Global definitions

loginPage    = 'https://apps.commercehub.com/account/login?service=https://dsm.commercehub.com/dsm/shiro-cas'
orderSummary = 'https://dsm.commercehub.com/dsm/gotoOrderSummary.do'

urlBase  = 'https://dsm.commercehub.com/dsm/gotoOrderRealmForm.do?'
openOrdersQuery = 'action=web_quickship&tabContext=web_quickship&status=open&substatus=no-activity&merchant='
invoicingQuery  = 'action=web_quickinvoice&tabContext=web_quickinvoice&merchant='



####------------------------------------------------------------------------------------------------- #
####------------------------------------- update merchants here ------------------------------------- #
tableBedBath = (0, 5, 3, 4)
tableShopHQ  = (3, 3, 1, 2)

merchantPids     = ['bedbath', 'snbc', 'belk'] # replace everything inside of square brackets
merchantNames    = {'bedbath': 'Bed Bath & Beyond', 'snbc': 'ShopHQ', 'belk': 'Belk'}
merchShipMeths   = {'bedbath': 'FEDX', 'snbc': 'UG', 'belk': 'FXSP'}
merchTableFormat = {'bedbath': tableBedBath, 'snbc': tableShopHQ}
#### ------------------------------------------------------------------------------------------------ #
#### ------------------------------------------------------------------------------------------------ #



browsers = {'Chrome': webdriver.Chrome, 'Edge': webdriver.Edge, 'Firefox': webdriver.Firefox}
browser  = None

trackingFile = 'tracking.csv'
trackingData = None


## ---------------------------------------------------------------------------------------------------------- #
# Class definitions

##  \class Error
#   \brief This class inherits the Enum class enabling enumerated error messages to be returned by
#          all functions within this script.
#
class Error(enum.Enum):
    ENone       = 0
    ELogin      = 1
    EDataImport = 2
    ENoOrders   = 3


## ---------------------------------------------------------------------------------------------------------- #
# Function definitions

##  \fn     AutoFillInvoice
#   \brief  This function locates and clicks the 'autofill' button associated with an order invoice.
#
def AutoFillInvoice(invoice, name):
    invoice.find_element_by_name(name).click()

##  \fn     Cleanup
#   \brief  This function closes any used resources and exits the script. This can be under normal
#           circumstances, or also if some error occured.
#
def Cleanup(result):
    if browser:
        browser.close()
    if result is not Error.ENone:
        print(f"An error occured. Error code {result} was issued.")
    sys.exit(0)

##  \fn     CloseTab
#   \brief  This function identifies the order stream active tab and clicks the link to close it.
#
def CloseTab():
    try:
        tab = browser.find_element_by_id('active-tab')
        x = tab.find_element_by_css_selector('a').click()
    except NoSuchElementException:
        pass

##  \fn     ExtractMerchants
#   \brief  This function extracts merchant id's available in the order summary page of OrderStream. This
#           can be done anytime there are open orders, but for obtaining new merchant id's make sure orders
#           exist for that merchant.
#   \return String
#
def ExtractMerchants():
    ids       = merchantPids
    linkTable = browser.find_element_by_class_name('windowpane')
    links     = linkTable.find_elements_by_css_selector('a')

    for link in links:
        href = link.get_attribute('href')
        mId  = href[href.rindex('=') + 1:]
        if mId not in ids:
            ids.append(mId)
    
    return "'" + "', '".join(ids) + "'"

##  \fn     ImportTrackingData
#   \brief  This function imports tracking data into a pandas dataframe. The global variable 'trackingData'
#           is used here and can later be referenced for individual cells.
#
def ImportTrackingData():
    global trackingData
    try:
        trackingData = pandas.read_csv(trackingFile, skip_blank_lines = True, index_col = 'Name')
    except FileNotFoundError:
        print("Woops! Did you forget the tracking file?")
        Cleanup(Error.EDataImport)
    
    if 'TrackingNumber' in trackingData.columns:
        trackingData = trackingData.T.to_dict()
        trackingData = {k.upper():v for k,v in trackingData.items()}
    else:
        Cleanup(Error.EDataImport)

##  \fn     Login
#   \brief  This function takes a username and password as input, as well as the preferred
#           selenium web driver to use, and logs in to the OrderStream website.
#   \return Error
#
def Login(username, password, driver):
    global browser
    
    webDriver = browsers.get(driver)
    browser = webDriver()
    browser.get(loginPage)
    
    userBox = browser.find_element_by_id('username')
    userBox.send_keys(username)

    passBox = browser.find_element_by_id('password')
    passBox.send_keys(password)

    browser.find_element_by_class_name('sign-in-button').click()
    time.sleep(1)

    try:
        browser.find_element_by_class_name('error-message')
        return Error.ELogin
    except NoSuchElementException:
        return Error.ENone

##  \fn     ProcessOpenOrders
#   \brief  This function opens the order invoice page for each merchant, fills in the required
#           information, also checking that multiple pages are filled, and finally submits the
#           completed information.
#   \return Error
#
def ProcessOpenOrders(merchant, noReview):
    print(f"Checking open orders for {merchantNames.get(merchant)}...", end='', flush=True)

    invoices = browser.find_elements_by_class_name("fw_widget_windowtag")

    errorTxt = 'No order(s) found that match the supplied criteria.'
    l = len(invoices)
    if l == 0 or (l == 1 and browser.find_element_by_class_name('fw_widget_windowtag_body').text == errorTxt):
        print('none found!', flush=True)
        CloseTab()
        return Error.ENoOrders

    try:
        pageSelector  = browser.find_element_by_class_name('fw_widget_pageselector_text').text
        pgSelTextList = [int(i) for i in pageSelector.split() if i.isdigit()]
        totalOrders   = pgSelTextList[2]
    except NoSuchElementException:
        return Error.ENoOrders

    reviewedOrders  = 0
    submittedOrders = 0
    while reviewedOrders < totalOrders:
        processedOrders = 0
        pageSize        = len(invoices)
        for invoice in invoices:
            ##  \brief Setup some initial variables with needed information. Some of the elements
            #          on the page are given a dynamic name/class/id based on the order number, 
            #          item number and finally the action, either tracking/shipping/shipped, etc.
            #
            orderNumber  = invoice.find_element_by_css_selector('a').get_attribute('href')
            orderNumber  = orderNumber[orderNumber.index('=') + 1:]
            customerName = invoice.find_element_by_class_name('framework_fiftyfifty_left_greenoutline').text
            customerName = customerName[9:customerName.index(',')].upper()
            trackingName = 'order(' + orderNumber + ').box(1).trackingnumber'
            shippingName = 'order(' + orderNumber + ').box(1).shippingmethod'
        
            ##  \brief If the customer name exists in the spreadsheet, process their order using
            #          the extracted information.
            #
            tn = trackingData.get(customerName)
            if tn is not None:
                ## 1) click autofill button for belk
                if merchant == 'belk':
                    btnAutoFillName = "autofill"
                    AutoFillInvoice(invoice, btnAutoFillName)

                ## 2) fill tracking number box
                tn = tn.get('TrackingNumber')
                trackingBox = invoice.find_element_by_name(trackingName)
                trackingBox.send_keys(tn)

                ## 3) fill shipping method dropdown box
                shippingDropdown = Select(invoice.find_element_by_id(shippingName))
                shippingDropdown.select_by_value(merchShipMeths.get(merchant))

                ## 4) extract remaining qty and fill shipped qty box, then extract item number
                orderTable = invoice.find_elements_by_class_name('or_numericdata')
                for j in range(4, len(orderTable), 4):
                    shipQty    = orderTable[j+1].text
                    elemId     = orderTable[j+1].get_attribute('id')
                    openPar    = elemId.rindex('(') + 1
                    closePar   = elemId.rindex(')')
                    itemNumber = elemId[openPar:closePar]

                    ## 5) setup element names using order number and item number
                    shipQtyName = 'order(' + orderNumber + ').box(1).item(' + itemNumber + ').shipped'
                    warName = 'order(' + orderNumber + ').box(1).item(' + itemNumber + ').shippingLineWarehouse'

                    ## 6) fill qty shipped box
                    shipQtyBox = invoice.find_element_by_name(shipQtyName)
                    shipQtyBox.send_keys(shipQty)

                    ## 7) fill in shipping warehouse (bedbath only)
                    if merchant == 'bedbath':
                        warehouse = Select(invoice.find_element_by_id(warName))
                        warehouse.select_by_value('LDM St Louis Park')
                
                # 8) increment processedOrders to indicate number of orders processed
                processedOrders += 1
            reviewedOrders += 1
        html = browser.find_element_by_tag_name('html')
        if processedOrders > 0:
            if noReview == False:
                input("Please check, does the information match? (hit enter when ready)")
            else:
                time.sleep(RandomTime(2, 3))
            submittedOrders += processedOrders
            browser.find_element_by_name('confirmbtn').click()
        elif pageSize < totalOrders and reviewedOrders < totalOrders:
                browser.find_element_by_class_name("fw_widget_pageselector_notcurrentpage").click()
        try:
            WebDriverWait(browser, 5).until(ec.staleness_of(html))
        except TimeoutException:
            pass
        invoices = browser.find_elements_by_class_name("fw_widget_windowtag")
    
    print('')
    print(f"Processed {submittedOrders} open order(s) for {merchantNames.get(merchant)}")
    CloseTab()

    return Error.ENone

##  \fn     ProcessInvoicing
#   \brief  This function opens the needs invoicing page for each merchant, fills in the required
#           information, also checking that multiple pages are filled, and finally submits the
#           completed information.
#   \return Error
#
def ProcessInvoicing(merchant):
    print(f"Checking invoices for {merchantNames.get(merchant)}...", end='', flush=True)

    invoices = browser.find_elements_by_class_name("fw_widget_windowtag")

    errorTxt = 'No order(s) found that match the supplied criteria.'
    l = len(invoices)
    if l == 0 or (l == 1 and browser.find_element_by_class_name('fw_widget_windowtag_body').text == errorTxt):
        print('none found!', flush=True)
        CloseTab()
        return Error.ENoOrders
    
    # debug section
    print(f'length of invoices: {len(invoices)}')
    print(errorTxt)
    print(browser.find_element_by_class_name('fw_widget_windowtag_body').text)

    try:
        pageSelector  = browser.find_element_by_class_name('fw_widget_pageselector_text').text
        pgSelTextList = [int(i) for i in pageSelector.split() if i.isdigit()]
        totalInvoices = pgSelTextList[2]
    except NoSuchElementException:
        print("got here!")
        return Error.ENoOrders

    processedInvoices = 0
    while processedInvoices < totalInvoices:
        tableFormat = merchTableFormat.get(merchant)
        start = tableFormat[0]
        step  = tableFormat[1]
        one   = tableFormat[2]
        two   = tableFormat[3]
        for invoice in invoices:
            ##  \brief Setup some initial variables with needed information. Some of the elements
            #          on the page are given a dynamic name/class/id based on the order number, 
            #          item number and finally the action, either tracking/shipping/shipped, etc.
            #
            orderNumber = invoice.find_element_by_css_selector('a').get_attribute('href')
            orderNumber = orderNumber[orderNumber.index('=') + 1:]

            ## 1) click the autofill button
            if merchant == "bedbath":
                btnAutoFillName = "autofill"
            elif merchant == "snbc":
                btnAutoFillName = "order(" + orderNumber + ").invoicenumber.autofill"
            AutoFillInvoice(invoice, btnAutoFillName)
            

            ## 2) extract remaining qty and fill shipped qty box, then extract item number
            invoiceTable = invoice.find_elements_by_class_name('or_numericdata')
            for j in range(start, len(invoiceTable), step):
                remainingQty = invoiceTable[j+one].text
                invoicedQty  = invoiceTable[j+two].find_element_by_xpath(".//input[@type='text']")
                invoicedQty.send_keys(remainingQty)

            processedInvoices += 1
        time.sleep(RandomTime(1, 4))
        html = browser.find_element_by_tag_name('html')
        browser.find_element_by_name('confirmbtn').click()
        WebDriverWait(browser, 5).until(ec.staleness_of(html))
        invoices = browser.find_elements_by_class_name("fw_widget_windowtag")
    
    print('')
    print(f"Invoiced {processedInvoices} order(s) for {merchantNames.get(merchant)}")
    CloseTab()

    return Error.ENone

##  \fn     RandomTime
#   \brief  This simple function returns a random number in the range 1-5
#   \return Integer
#
def RandomTime(start = 1, finish = 6):
    return random.randrange(start, finish)




## ---------------------------------------------------------------------------------------------------------- #
# Main function

def Main():
    """ \brief Argparse section

    """
    argParseDescription = ('OrderStream automation utility. Supply your username and password at the command '
                           'line, and let the script do the rest! Also available is the ability to extract '
                           'merchant IDs for all open orders. This is necessary when new merchants are '
                           'added. The formatted output can then be pasted into the script.')
    userHelp            = 'Enter your OrderStream username. (required)'
    passHelp            = 'Enter your OrderStream password. (required)'
    browserHelp         = ('Enter your preferred browser (optional). Choices are "Chrome", "Edge", and '
                          '"Firefox", default is Chrome.')
    extractHelp         = ('Include this flag if you wish to only extract the merchant IDs of open orders. '
                           'This will give string output in a formatted list that you can paste inside '
                           'the square brackets of the variable "merchantPids" within the script.')
    noreviewHelp        = ('Include this flag if you wish to run the script without manual review for '
                           'processing open orders. This will cause the script to skip the pause after '
                           'tracking and other info is filled in.')
    
    parser = argparse.ArgumentParser(description = argParseDescription)
    
    parser.add_argument('-u', required = True, dest = 'userName', metavar = 'username', help = userHelp)
    parser.add_argument('-p', required = True, dest = 'passWord', metavar = 'password', help = passHelp)
    parser.add_argument('-b', dest = 'browser', metavar = 'browser', default = 'Chrome', help = browserHelp)
    parser.add_argument('--extract', action = 'store_true', help = extractHelp)
    parser.add_argument('--noreview', action = 'store_true', help = noreviewHelp)
    
    args = parser.parse_args()

    
    """ \brief Main section

    """
    ## 1) import current order tracking info
    ImportTrackingData()

    ## 2) attempt to login, quit if unsuccessful
    if Login(args.userName, args.passWord, args.browser) == Error.ELogin:
        Cleanup(Error.ELogin)
    
    ## 3) if extract flag is given, go to order summary and pull merchant IDs
    if args.extract:
        browser.get(orderSummary)
        print('The list of merchant IDs for open orders is as follows:', ExtractMerchants())
        Cleanup(Error.ENone)

    ## 4) finally, process orders and invoicing
    for merchant in merchantPids:
        browser.get(urlBase + openOrdersQuery + merchant)
        ProcessOpenOrders(merchant, args.noreview)
        if merchant != 'belk':
            browser.get(urlBase + invoicingQuery + merchant)
            ProcessInvoicing(merchant)
    
    ## 5) print completion message
    print("Completed processing open orders and/or invoicing!")

if __name__ == '__main__':
    Main()

