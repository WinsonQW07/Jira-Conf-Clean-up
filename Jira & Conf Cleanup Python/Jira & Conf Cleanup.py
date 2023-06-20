# -*- coding: utf-8 -*-
"""
Spyder Editor

"""

import csv
import pandas as pd
import numpy as np

from datetime import date
from datetime import datetime

#EXAMPLE OF USER INPUTTED LOCAL PATH (- #)
#C:\Users\ww459\Desktop\Jira & Conf Cleanup Python\Atlassian Users 2023-06-09.csv

#Returns user inputted local path as a string
path_file = input("Please enter the local path to the CSV file. Don't include quotation marks: ")

#turns the user inputted string into a raw string
raw_path_file = repr(path_file)[1:-1]
    
#reads csv file using the raw string local path
user_data = pd.read_csv(raw_path_file, delimiter = ",")

#xlsx file to write to
sorted_users = "Sorted_Users_xlsx.xlsx"

print(user_data.columns)

#different lists of users
all_users = [[row[col] for col in user_data.columns] for row in user_data.to_dict('records')]
users_not_access_jira_and_conf = []
users_not_access_jira = []
users_not_access_conf = []
users_no_jira_over_one_year = []
users_no_conf_over_one_year = []

#adds column names to the lists
users_not_access_jira_and_conf.append(user_data.columns)
users_not_access_jira.append(user_data.columns)
users_not_access_conf.append(user_data.columns)
users_no_jira_over_one_year.append(user_data.columns)
users_no_conf_over_one_year.append(user_data.columns)

#function that prints to both the console and into the text file
def write_line(string):
    print(string)
    # conv_string = string
    # with open("User Data.txt", "a") as o:
    #     #checks if parameter is a series
    #     if(isinstance(string, pd.Series)):
    #         conv_string = string.to_string()
    #     o.write(conv_string)
        
#function for checking if a user has never accessed either Jira or Confluence
def check_last_login(last_login, user_list, column_number):
    for i in range(0, len(user_data)):
        if(last_login[i][column_number] == "Never accessed"):
            user_list.append(all_users[i])

#loops through all users to find those who have never accessed jira AND confluence and adds them to its respective list
for i in range(0, len(user_data)):
    if(all_users[i][6] == "Never accessed" and all_users[i][7] == "Never accessed"):
        users_not_access_jira_and_conf.append(all_users[i])
        
#writes list of users to csv file
# with open(sorted_users, "w") as csvfile:
#     csvwriter = csv.writer(csvfile)
    
#     csvwriter.writerow(user_data.columns)

#     csvwriter.writerows(users_not_access_jira_and_conf)

#loops through all users to find those who have never used Jira
check_last_login(all_users, users_not_access_jira, 6)
        
#loops through all users to find those who have never used Confluence
check_last_login(all_users, users_not_access_conf, 7)

current_date = date.today()

#finds the users who have not used Jira in over a year
for i in range(0, len(user_data)):
    
    #skips users who have not logged in before
    if(all_users[i][6] == "Never accessed"):
        continue
    
    #skips rows with no values
    if(pd.isna(all_users[i][0])):
        continue
    
    #turns the date string into a datetime.date object
    date_object = datetime.strptime(all_users[i][6], "%d-%b-%y").date()
    
    #finds the number of days since the user last logged in compared to the current date
    time_difference = current_date - date_object
    
    #gets the users who have not used Jira in over a year
    if abs(time_difference.days) > 365:
        users_no_jira_over_one_year.append(all_users[i])

#finds the users who have not used Confluence in over a year
for i in range(0, len(user_data)):
    
    #skips users who have not logged in before
    if(all_users[i][7] == "Never accessed"):
        continue
    
    #turns the date string into a datetime.date object
    date_object = datetime.strptime(all_users[i][7], "%d-%b-%y").date()
    
    #finds the number of days since the user last logged in compared to the current date
    time_difference = current_date - date_object
    
    #gets the users who have not used Jira in over a year
    if abs(time_difference.days) > 365:
        users_no_conf_over_one_year.append(all_users[i])
        
#creates a dataframe for each list of users
df1 = pd.DataFrame(users_not_access_jira_and_conf)
df2 = pd.DataFrame(users_not_access_jira)
df3 = pd.DataFrame(users_not_access_conf)
df4 = pd.DataFrame(users_no_jira_over_one_year)
df5 = pd.DataFrame(users_no_conf_over_one_year)

#writes to the xlsx file in multiple sheets
with pd.ExcelWriter(sorted_users) as engine:
    df1.to_excel(excel_writer=engine, sheet_name="JIRA & CONF Not Accessed")
    df2.to_excel(excel_writer=engine, sheet_name="JIRA Not Accessed")
    df3.to_excel(excel_writer=engine, sheet_name="CONF Not Accessed")
    df4.to_excel(excel_writer=engine, sheet_name="Jira over 1 year not accessed")
    df5.to_excel(excel_writer=engine, sheet_name="CONF over 1 year not accessed")