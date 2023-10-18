
import os
import shutil
import xlwings as xw
import datetime as dt

day_of_week = input("What day of the week is it? ")

if (day_of_week == "Monday"):
    date = input("What is the date of the file? ")
    format_name = "C:/Users/Asset Management - Documents/Vacancy Case Studies/Vacancy Focus Properties (" + date + ") AM Comments.xlsm" 

    VCS = xw.Book(format_name, update_links= False)

    sheet = VCS.sheets['UA']

    date1= input("What is the date for cell C1? ")
    #changes date in cell C1
    sheet.range("C1").value = date1

    sheet.api.AutoFilterMode = False
    #clears column A to Q for UA tab in case studies file
    sheet.range("A3:Q5500").clear_contents()

    td_date = input("What is today's date? ")

    UA_name = "U:/CDNMF/Reports/Unit Availability/2022/" + "Unit Availability " + td_date + ".xlsm"
    UA = xw.Book(UA_name, update_links=False)

    new_sheet = UA.sheets['Availability']
    new_sheet.api.AutoFilterMode = False

    #copys sheet to VCS report
    new_sheet.copy(after=VCS.sheets[0])

    #stores values in column A to Q
    my_values = UA.sheets['Availability'].range("A3:Q5500").options(ndim=2).value

    #copies column A to Q
    VCS.sheets["UA"].range("A3:Q5500").value = my_values

    new_sheet.delete()

    sheet.range("F3:G3800").api.Replace("", "-")
    sheet.range("L3:L3800").api.Replace("", "-")

    old_date = input("What is the current Old UA (2 weeks back) m-dd-yy? ")

    for sheet in VCS.sheets:
        
        print("Currently on " + sheet)
        #links current UA
        UA.sheets[sheet].range("A1:CC5500").api.Replace(date, td_date)
        #links last weeks UA
        UA.sheets[sheet].range("A1:CC5500").api.Replace(old_date,date)

    print("\nRelinking finished\n")

if (day_of_week == "Tuesday"):
    
    date = input("What is the date of the file (Sharepoint, last week monday)? m-dd-yy ")
    format_name = "C:/Users/Asset Management - Documents/Vacancy Case Studies/Vacancy Focus Properties (" + date + ") AM Comments.xlsm"
    
    td_file_date = input("What is the current date (Monday date)? m-dd-yy ")
    dest_dir = "U:/CDNMF/Reports/Vacancy-Case Studies/DD - HC/" + "Vacancy Focus Properties-HC (" + td_file_date + ").xlsm"
    
    #copys file under our new name
    shutil.copy(format_name, dest_dir)
    
    print("copy complete")
    
    VCS = xw.Book(dest_dir, update_links= False)
    
    DD_folder_date = input("What is the folder date (Monday date)? YYYY-mm-dd ")
    td_file_date_DD = input("What is the current date (Monday date)? mm-dd-yy ")

    DD_name = "U:/Leasing Data/DD - Semi-Monthly Update/" + DD_folder_date + "/DD Weekly Update " + td_file_date_DD + ".xlsx"
    DD = xw.Book(DD_name,update_links= False)
    
    #copys DD sheets to VCS report
    DD.sheets["Vacancy-BC"].copy(before=VCS.sheets[0])
    DD.sheets["Vacancy - CDNMF"].copy(before=VCS.sheets[0])
    
    #all data is here now, we will now unhide/hide sheets, and check numbers.

if (day_of_week == "Wednesday"):
    
    date = input("What is the date of the file (Sharepoint)? m-dd-yy ")
    format_name = "C:/Users/Asset Management - Documents/Vacancy Case Studies/Vacancy Focus Properties (" + date + ") AM Comments.xlsm"
    
    td_date = input("What is the current date (Monday)? m-dd-yy ") 
    td_file_date = input("What is the current date (Monday)? mm-dd-yy ")
    dest_dir = "U:/CDNMF/Reports/Vacancy-Case Studies/AM Sharepoint/" + "Vacancy Focus Properties (" + td_date + ") AM Comments.xlsm"
    
    #copys file under our new name
    shutil.copy(format_name, dest_dir)
    
    VCS = xw.Book(dest_dir, update_links= False)
    
    DD_folder_date = input("What is the folder date (Monday Date)? YYYY-mm-dd ")

    DD_name = "U:/Leasing Data/DD - Semi-Monthly Update/" + DD_folder_date + "/DD Weekly Update " + td_file_date + ".xlsx"
    DD = xw.Book(DD_name,update_links= False)
       
    
    
    
    
    
    
