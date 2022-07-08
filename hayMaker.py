#! python3
# hayMaker.py - A rain maker automation program.

##Importing Modules####
import sys
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

# define the scope
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('C:\\Users\\IT\\Downloads\\haymaker-332905-a02bef8aacbd.json', scope)

# authorize the clientsheet 
client = gspread.authorize(creds)




dictSelectA = {"Network Domain": "#select2-NetworkDomain-container", "Region":"#select2-RegionID-container","City": "#select2-CityID-container > span" ,"Issue On":"#select2-IssueOn-container > span",
              "Mode of Ticket":"#select2-ModeOfTicketID-container","Ticket Status":"#select2-TicketStatusID-container > span","Fault":"#select2-Fault-container > span"
               ,"Impact Area":"#select2-ImpactType-container" ,"Impact":"#select2-ImpactID-container > span","Ticket Owner":"#select2-TicketOwnerEMPID-container"
               ,"Escalated Person":"#select2-EscalatedPersonEMPID-container > span","Node A":"#select2-NodeAID-container > span",  }

dictSelectB = {"Region B" : "#select2-RegionBID-container",
                 "City B": "#select2-CityBID-container", "Node B": "#select2-NodeBID-container"}

ListDictSelectA = list(dictSelectA.values())
ListDictSelectB = list(dictSelectB.values())


# dictTicketA = {"Network Domain": "TXN","Region": "North","City": "Islamabad","Issue On": "Connectivity", "Mode of Ticket": "Longhaul - LH", "Ticket Status": "Down", "Fault": "Connectivity Down"
#               , "Impact Area": "Core", "Impact": "Non Service Affecting", "Ticket Owner": "EMP00684 Muhammad Irfan Khan Durani",
#               "Escalated Person": "EMP00684 Muhammad Irfan Khan Durani", "Node A": "R4/ISB", }
#
# dictTicketB = {"Region B" : "Central",
#                  "City B": "Talagang", "Node B": "R4/TLG"}




dictChooseA = {"Network Domain":'//*[@id="kt_body"]/span/span/span[1]/input',"Region":'//*[@id="kt_body"]/span/span/span[1]/input',"City":'//*[@id="kt_body"]/span/span/span[1]/input'
             ,"Issue On":'//*[@id="kt_body"]/span/span/span[1]/input',"Mode of Ticket":'//*[@id="kt_body"]/span/span/span[1]/input'
            ,"Ticket Status":'//*[@id="kt_body"]/span/span/span[1]/input'
            ,"Fault":'//*[@id="kt_body"]/span/span/span[1]/input',"Impact Area":'//*[@id="kt_body"]/span/span/span[1]/input'
            ,"Impact":'//*[@id="kt_body"]/span/span/span[1]/input',"Ticket Owner":'//*[@id="kt_body"]/span/span/span[1]/input',"Escalated Person":'//*[@id="kt_body"]/span/span/span[1]/input'
            ,"Node A":'//*[@id="kt_body"]/span/span/span[1]/input',  }

dictChooseB = {"Region B" : '//*[@id="kt_body"]/span/span/span[1]/input',
                 "City B": '//*[@id="kt_body"]/span/span/span[1]/input', "Node B": '//*[@id="kt_body"]/span/span/span[1]/input'}

ListDictChooseA = list(dictChooseA.values())
ListDictChooseB = list(dictChooseB.values())

#################Google Sheets Data#######################


# get the instance of the Spreadsheet
sheet = client.open('Data')

# get the first sheet of the Spreadsheet
sheet1 = sheet.get_worksheet(0)
dictTicketA = {}
dictTicketB = {}

for rowValue in range (1,23):
    if rowValue <= 12:
        dictTicketA.update({sheet1.cell(col = 1, row = rowValue ).value:sheet1.cell(col = 2, row=rowValue ).value})
    else:
        dictTicketB.update({sheet1.cell(col = 1, row = rowValue).value:sheet1.cell(col = 2, row=rowValue).value})

# for columnValue in range (1,17):
#     if columnValue <= 12:
#         dictTicketA.update({sheet1.cell(row = 1, column=columnValue).value:sheet1.cell(row=2, column=columnValue).value})
#     else:
#         dictTicketB.update({sheet1.cell(row = 1, column=columnValue).value:sheet1.cell(row=2, column=columnValue).value})

ListDictTicketA = list(dictTicketA.values())
ListDictTicketB = list(dictTicketB.values())

#web = webdriver.Chrome(executable_path=r'C:\Users\GSAC\Desktop\pyChapterProjects\chromedriver_win32\chromedriver.exe')
web = webdriver.Chrome()
web.get('https://noc.rainmaker.pk/Ticket/Generate')

time.sleep(1)

# userElem = web.find_element_by_xpath('//*[@id="UserName"]')
# userElem.send_keys('muhammad.saad@multinet.com.pk')

# passwordElem = web.find_element_by_xpath('//*[@id="Password"]')
# passwordElem.send_keys('Multi@@2013')

userElem = web.find_element_by_xpath('//*[@id="UserName"]')
userElem.send_keys(ListDictTicketB[8])

passwordElem = web.find_element_by_xpath('//*[@id="Password"]')
passwordElem.send_keys(ListDictTicketB[9])

submit = web.find_element_by_xpath('//*[@id="kt_login_singin_form_submit_button"]')
submit.click()

#################Select Region####################
time.sleep(1)


if ListDictTicketB[6] == 'Medium':
    priority = '#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(1) > div > div > div > div:nth-child(12) > div > label.radio.radio-warning > span'
    selectPriority = web.find_element_by_css_selector(priority)
    selectPriority.click()
else:
    priority = '#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(1) > div > div > div > div:nth-child(12) > div > label.radio.radio-danger > span'
    selectPriority = web.find_element_by_css_selector(priority)
    selectPriority.click()


if (ListDictTicketA[4] == 'Longhaul - LH' or ListDictTicketA[4] == 'Long Haul - EXT - EX'):
    #if ListDictTicketA[6] == ('Connectivity Down' or 'Fiber'):
    if (ListDictTicketA[6] == 'Connectivity Down' or ListDictTicketA[6] == 'Fiber'):
        print('inside correct loop')
        for i in range (0,12):
            selectText = web.find_element_by_css_selector(ListDictSelectA[i])
            selectText.click()

            TicketValue = ListDictTicketA[i]
            chooseText = web.find_element_by_xpath(ListDictChooseA[i])
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)
            time.sleep(0.5)

        diffRegion = '#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(2) > div > div.col-md-12.form-group > div > label > span'
        print("inside selector")
        SelectDiffRegion = web.find_element_by_css_selector(diffRegion)
        SelectDiffRegion.click()

        for j in range (0,3):
            selectText = web.find_element_by_css_selector(ListDictSelectB[j])
            selectText.click()

            TicketValue = ListDictTicketB[j]
            chooseText = web.find_element_by_xpath(ListDictChooseB[j])
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)
            time.sleep(0.5)

    elif (ListDictTicketA[6] == 'Power' or ListDictTicketA[6] == 'System'):
        print('inside power loop')
        for i in range(0, 12):
            selectText = web.find_element_by_css_selector(ListDictSelectA[i])
            selectText.click()

            TicketValue = ListDictTicketA[i]
            chooseText = web.find_element_by_xpath(ListDictChooseA[i])
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)
            time.sleep(0.5)
            
if ListDictTicketB[5] == '10G':
        Lambda = '#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(5) > div > div > div > div.col-md-4.txnField > div > label:nth-child(1) > span'
        SelectLambda = web.find_element_by_css_selector(Lambda)
        SelectLambda.click()
elif ListDictTicketB[5] == '2.5G':
        Lambda = '#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(5) > div > div > div > div.col-md-4.txnField > div > label:nth-child(2) > span'
        SelectLambda = web.find_element_by_css_selector(Lambda)
        SelectLambda.click()




elif ListDictTicketA[4] == ('Metronet - MT'):

    if (ListDictTicketA[6] == 'Connectivity Down' or ListDictTicketA[6] == 'Fiber'):
        for i in range (0,12):
            selectText = web.find_element_by_css_selector(ListDictSelectA[i])
            selectText.click()

            TicketValue = ListDictTicketA[i]
            chooseText = web.find_element_by_xpath(ListDictChooseA[i])
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)
            time.sleep(0.5)

        if ListDictTicketB[7] == 'Leg B':
            trueRedundant= web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-2 > div > label > span')
            trueRedundant.click()

            SelectRedundant = web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-10 > div > label:nth-child(2) > span')
            SelectRedundant.click()
            
             ######selecting Node B metro################

            selectText = web.find_element_by_css_selector(ListDictSelectB[2])
            selectText.click()

            TicketValue = ListDictTicketB[2]
            chooseText = web.find_element_by_xpath('//*[@id="kt_body"]/span/span/span[1]/input')
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)

        elif ListDictTicketB[7] == 'Leg A':
            trueRedundant= web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-2 > div > label > span')
            trueRedundant.click()

            SelectRedundant = web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-10 > div > label:nth-child(1) > span')
            SelectRedundant.click()
            
             ######selecting Node B metro################

            selectText = web.find_element_by_css_selector(ListDictSelectB[2])
            selectText.click()

            TicketValue = ListDictTicketB[2]
            chooseText = web.find_element_by_xpath('//*[@id="kt_body"]/span/span/span[1]/input')
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)
    
    elif (ListDictTicketA[6] == 'Power' or ListDictTicketA[6] == 'System'):
        print('inside power loop')
        for i in range(0, 12):
            selectText = web.find_element_by_css_selector(ListDictSelectA[i])
            selectText.click()

            TicketValue = ListDictTicketA[i]
            chooseText = web.find_element_by_xpath(ListDictChooseA[i])
            chooseText.send_keys(TicketValue)
            chooseText.send_keys(Keys.RETURN)
            time.sleep(0.5)
    
elif ListDictTicketA[4] == ('Cross Border - CB'):
    for i in range (0,12):
        selectText = web.find_element_by_css_selector(ListDictSelectA[i])
        selectText.click()

        TicketValue = ListDictTicketA[i]
        chooseText = web.find_element_by_xpath(ListDictChooseA[i])
        chooseText.send_keys(TicketValue)
        chooseText.send_keys(Keys.RETURN)
        time.sleep(0.5)

    diffRegion = '#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(2) > div > div.col-md-12.form-group > div > label > span'
    SelectDiffRegion = web.find_element_by_css_selector(diffRegion)
    SelectDiffRegion.click()

    for j in range(0, 3):
        selectText = web.find_element_by_css_selector(ListDictSelectB[j])
        selectText.click()

        TicketValue = ListDictTicketB[j]
        chooseText = web.find_element_by_xpath(ListDictChooseB[j])
        chooseText.send_keys(TicketValue)
        chooseText.send_keys(Keys.RETURN)

    if ListDictTicketB[7] == 'Leg B':
        trueRedundant= web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-2 > div > label > span')
        trueRedundant.click()

        SelectRedundant = web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-10 > div > label:nth-child(2) > span')
        SelectRedundant.click()

    elif ListDictTicketB[7] == 'Leg A':
        trueRedundant= web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-2 > div > label > span')
        trueRedundant.click()

        SelectRedundant = web.find_element_by_css_selector('#kt_page_sticky_card > div.card-body > div:nth-child(1) > div:nth-child(3) > div > div > div > div:nth-child(1) > div > div.col-md-12.form-group > div > div.col-md-10 > div > label:nth-child(1) > span')
        SelectRedundant.click()
######selecting Node B CB################

    selectText = web.find_element_by_css_selector(ListDictSelectB[2])
    selectText.click()

    TicketValue = ListDictTicketB[2]
    chooseText = web.find_element_by_xpath('//*[@id="kt_body"]/span/span/span[1]/input')
    chooseText.send_keys(TicketValue)
    chooseText.send_keys(Keys.RETURN)


commValue = ListDictTicketB[3]
choosecommValue = web.find_element_by_xpath('//*[@id="Remarks"]')
choosecommValue.send_keys(commValue)



##########Generating Mail Alert ###########################


if ListDictTicketA[4] == ('Longhaul - LH') or ListDictTicketA[4] == ('Long Haul - EXT - EX') or ListDictTicketA[4] ==('Cross Border - CB'):
    if ListDictTicketA[6] == ('Connectivity Down' or 'Fiber'):
        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(1) > div > label:nth-child(6) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()

    elif  ListDictTicketA[6] == 'Power':
        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(5) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()
    else:
        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(6) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()

elif ListDictTicketA[4] == 'Metronet - MT':
    if ListDictTicketA[6] == ('Connectivity Down' or 'Fiber') and (ListDictTicketA[1] == 'Central'):
        mailAlert1 = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(4) > span'
        SelectmailAlert1 = web.find_element_by_css_selector(mailAlert1)
        SelectmailAlert1.click()

        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(6) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()

    elif ListDictTicketA[6] == ('Connectivity Down' or 'Fiber') and (ListDictTicketA[1] == 'South'):
        mailAlert1 = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(2) > span'
        SelectmailAlert1 = web.find_element_by_css_selector(mailAlert1)
        SelectmailAlert1.click()

        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(6) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()

    elif ListDictTicketA[6] == ('Connectivity Down' or 'Fiber') and (ListDictTicketA[1] == 'North'):
        mailAlert1= '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(3) > span'
        SelectmailAlert1 = web.find_element_by_css_selector(mailAlert1)
        SelectmailAlert1.click()


        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(6) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()

    elif ListDictTicketA[6] == 'Power':
        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(5) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()
    else:
        mailAlert = '#kt_page_sticky_card > div.card-body > div:nth-child(2) > div > div > div > div.row > div:nth-child(2) > div > label:nth-child(6) > span'
        SelectmailAlert = web.find_element_by_css_selector(mailAlert)
        SelectmailAlert.click()
print(ListDictTicketA[6])

selectSmsAlert = web.find_element_by_css_selector('#select2-groupId-container')
selectSmsAlert.click()

smsValue = 'TXN Broadcast'
chooseSMS = web.find_element_by_xpath('//*[@id="kt_body"]/span/span/span[1]/input')
chooseSMS.send_keys(smsValue)
chooseSMS.send_keys(Keys.RETURN)

while(True):
    pass
