#! /usr/bin/env python3

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import argparse
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import re

'''
Gspread docs: https://gspread.readthedocs.io/en/latest/user-guide.html#selecting-a-worksheet
'''
#API Constants
scope = ["https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

#Spreadsheet Constants
userInput = client.open('Finances').get_worksheet(0)
endOfSummer = client.open('Finances').get_worksheet(2)
budget = client.open('Finances').get_worksheet(1)
expenseCategories = ['Groceries','Eating Out','Frivolous','Insurance','House Payment', 'Investment', 'Gas', 'Car Payment', 'Correction', 'Donation','Health']
incomeCategories = ['Work','VA','School','Present','Interest','Refund','Correction']

def getDate():
    return str(datetime.date(datetime.now()))

def addExpense (EorI, Name, Amount, Category):
    match = False
    for title in expenseCategories:
        if Category.lower() in title.lower():
            match = True
            Category = title
            break
    if not match:
        print('Incorrect Category\n  Options include:')
        print(expenseCategories)
        return
    insertLine = len(endOfSummer.col_values(2)) + 1
    endOfSummer.insert_row(None,index=insertLine)
    endOfSummer.update_cell(insertLine,2,getDate())
    endOfSummer.update_cell(insertLine,3,Name)
    endOfSummer.update_cell(insertLine,4,float(Amount))
    endOfSummer.update_cell(insertLine,5,Category)

def addIncome (EorI, Name, Amount, Category):
    match = False
    for title in incomeCategories:
        if Category.lower() in title.lower():
            match = True
            Category = title
            break
    if not match:
        print('Incorrect Category\n  Options include:')
        print(incomeCategories)
        return
    insertLine = len(endOfSummer.col_values(7)) + 1
    endOfSummer.update_cell(insertLine,7,getDate())
    endOfSummer.update_cell(insertLine,8,Name)
    endOfSummer.update_cell(insertLine,9,float(Amount))
    endOfSummer.update_cell(insertLine,10,Category)

def scrape():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--no-sandbox')
    driver = webdriver.Chrome('/usr/bin/chromedriver', chrome_options=chrome_options)
    driver.maximize_window()
    driver.get("https://www.rcu.org/")
    # assert "rcu" in driver.title
    elem = driver.find_element_by_xpath('/html/body/header/div/div[1]/div/nav/div[1]/button[2]')
    time.sleep(2)
    elem.click()
    # time.sleep(2)
    elem = driver.find_element_by_id("Name")
    elem.clear()
    elem.send_keys("pashbylogan")
    time.sleep(2)
    elem = driver.find_element_by_id("Password")
    elem.clear()
    elem.send_keys("nasa-12$-tesla")
    elem.send_keys(Keys.ENTER)
    time.sleep(2)
    assert "No results found." not in driver.page_source
    driver.find_element_by_link_text('Logan\'s Checking').click()
    time.sleep(100)
    driver.close()

def updateBudget():
    dateVals = endOfSummer.col_values(2)[1:]
    priceVals = endOfSummer.col_values(4)[1:]
    priceVals = [float(i) for i in priceVals]
    categoryVals = endOfSummer.col_values(5)[1:]
    categoryIdx = []
    insuranceSum,houseSum,frivolousSum,foodSum,investSum,carSum,healthSum,donationSum = 0,0,0,0,0,0,0,0
    for i in range(len(dateVals)):
        if str(int(getDate()[5:7])) in re.split('\-|\/',str(dateVals[i]))[0] and str(int(getDate()[0:4])) == re.split('\-|\/',str(dateVals[i]))[2]:
            categoryIdx.append(i)
    for z in categoryIdx:
        if not categoryVals[z] == '':
            if categoryVals[z].lower() == 'insurance':
                insuranceSum += priceVals[z]
            elif categoryVals[z].lower() == 'eating out':
                foodSum += priceVals[z]
            elif categoryVals[z].lower() == 'frivolous':
                frivolousSum += priceVals[z]
            elif categoryVals[z].lower() == 'groceries':
                foodSum += priceVals[z]
            elif categoryVals[z].lower() in 'house payment':
                houseSum += priceVals[z]
            elif categoryVals[z].lower() == 'investment':
                investSum += priceVals[z]
            elif categoryVals[z].lower() in 'health':
                healthSum += priceVals[z]
            elif categoryVals[z].lower() == 'gas':
                carSum += priceVals[z]
            elif categoryVals[z].lower() in 'car payment':
                carSum += priceVals[z]
            elif categoryVals[z].lower() == 'donation':
                donationSum += priceVals[z]
            else:
                None
    budget.update('B2',115-insuranceSum)
    budget.update('B3',250-houseSum)
    budget.update('B4',200-frivolousSum)
    budget.update('B5',200-foodSum)
    budget.update('B6',100-investSum)
    budget.update('B7',20-carSum)
    budget.update('B8',0-healthSum)
    budget.update('B9',0-donationSum)

def updateIncome():
    dateVals = endOfSummer.col_values(7)[1:]
    priceVals = endOfSummer.col_values(9)[1:]
    priceVals = [float(i) for i in priceVals]
    categoryVals = endOfSummer.col_values(10)[1:]
    categoryIdx = []
    workSum,vaSum, schoolSum, presentSum, interestSum,refundSum = 0,0,0,0,0,0
    for i in range(len(dateVals)):
        if str(int(getDate()[5:7])) in re.split('\-|\/',str(dateVals[i]))[0] and str(int(getDate()[0:4])) == re.split('\-|\/',str(dateVals[i]))[2]:
            categoryIdx.append(i)
    for z in categoryIdx:
        if not categoryVals[z] == '':
            if categoryVals[z].lower() == 'work':
                workSum += priceVals[z]
            elif categoryVals[z].lower() == 'va':
                vaSum += priceVals[z]
            elif categoryVals[z].lower() == 'school':
                schoolSum += priceVals[z]
            elif categoryVals[z].lower() == 'present':
                presentSum += priceVals[z]
            elif categoryVals[z].lower() == 'interest':
                interestSum += priceVals[z]
            elif categoryVals[z].lower() == 'refund':
                refundSum += priceVals[z]
            else:
                None

    budget.update('D2',1000-vaSum)
    budget.update('D3',300-workSum)
    # budget.update('D4',2000-schoolSum)

def main():
    #parse arguments
    # args = parse_all_args()

    EorI = input('Expense (e) or Income (i): ')
    Name = input('Name of money flow: ')
    Amount = input('Amount: ')
    Category = input('Now choose the category of money flow. Would you like to see the categories?(y/n) ')
    if Category.lower() == 'y':
        print('The categories are:')
        if EorI.lower() == 'e':
            print(expenseCategories)
            Category = input('Category of money flow: ')
        else:
            print(incomeCategories)
            Category = input('Category of money flow: ')
    else:
        Category = input('Category of money flow: ')

    if EorI.lower() == 'e':
        addExpense(EorI, Name, Amount, Category)
    elif EorI.lower() == 'i':
        addIncome(EorI, Name, Amount, Category)
    else:
        print('Incorrect E or I')
    
    updateBudget()
    updateIncome()
    print('Finance Update Complete')
    print('Your current balance is:',userInput.acell('C6').value)


if __name__ == '__main__':
    main()
