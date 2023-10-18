import os
import shutil
import xlwings as xw
import datetime as dt


day_of_week = input("What is the day today ? Monday or Tuesday? ")

#this input helps us track where we are in the week
on_cycle = input("If you are on-cycle enter 0, if not please enter the number of days missed since Monday ")

#this loop ensures a value is entered
while (on_cycle == ""):
    on_cycle = input("You have not entered a value, please try again ")

#we make this a numebr so we can use it later
on_cycle = int(on_cycle) * -1

if (day_of_week == "Monday"):    # defining the destination directory
    #/python tester folder
    dest_dir = "U:/" + "/python tester folder"

    #now listdir will listen to dest_dir
    os.chdir(dest_dir)

    today_date = dt.date.today()  + dt.timedelta(on_cycle)
    last_week_date = today_date - dt.timedelta(7)

    #collects old file name from user
    old_file_name = dt.date.strftime(last_week_date,'%#m-%d-%y')
    #formats old file name
    formated_old_file_name = "Unit Availability " + old_file_name + ".xlsm"

    #collects new file name from user
    new_file_name = dt.date.strftime(today_date,'%#m-%d-%y')
    #fomats new file name
    formated_new_file_name = "Unit Availability " + new_file_name + ".xlsm"


    #changes file path to give copied file new name
    dest_dir = dest_dir + "/" + formated_new_file_name

    #copys file under our new name
    shutil.copy(formated_old_file_name, dest_dir)

    #
    # Now we change the date in cell A3, and copy paste the availability and Suite Reno downloads
    #

    #opens this weeks UA
    UA = xw.Book(formated_new_file_name, update_links=False)
    #goes to summary sheet
    sheet = UA.sheets['Summary']

    #changes date in cell A3
    sheet.range("A3").value = today_date

    #switches to availability sheet
    sheet = UA.sheets["Availability"]

    sheet.api.AutoFilterMode = False

    #clears column A to M for availability
    sheet.range("A2:M8000").clear_contents()

    download_date = dt.date.strftime(today_date,'%#m-%#d-%y')

    #adjusts file name
    UA_download = "U:/CDNMF/Unit Availability Download/" + download_date + ".xls"

    #opens UA download
    download = xw.Book(UA_download)

    download.sheets[0].api.AutoFilterMode = False

    #copys sheet to UA report
    download.sheets[0].copy(after=UA.sheets[0])


    UAlen = UA.sheets["report1"].range('A7').end('down').row

    #stores values in column A to C
    my_values = UA.sheets["report1"].range("A7:C" + str(UAlen)).options(ndim=2).value

    #copies column A to C
    UA.sheets["Availability"].range('A2:C8000').value = my_values

    #stores values in column F to O
    my_values = UA.sheets["report1"].range("F7:O" + str(UAlen)).options(ndim=2).value

    #copies column F to O
    UA.sheets["Availability"].range('D2:M8000').value = my_values

    sheet = UA.sheets['report1']
    sheet.delete()

    download.close()

    #clears column A to X for Renno
    UA.sheets["Availability"].range("N3:AT8000").clear_contents()

    formula = UA.sheets["Availability"].range("N2:AT2").options(ndim=2).formula

    #copies column A to C
    UA.sheets["Availability"].range("N2:AT" + str(UAlen)).formula = formula

    #switches to Reno sheet
    sheet = UA.sheets["Reno"]

    sheet.api.AutoFilterMode = False

    #clears column A to X for Renno
    sheet.range("A2:X55000").clear_contents()

    sheet.range("Y3:AB55000").clear_contents()

    #formats suite reno file name
    Suitereno_name = "U:/CDNMF/Suite Reno Download/" + download_date + ".xls"

    #opens suitereno file
    Suitereno = xw.Book(Suitereno_name)

    Suitereno.sheets[0].api.AutoFilterMode = False

    #copys sheet to UA report
    Suitereno.sheets[0].copy(after=UA.sheets[0])

    Renolen = UA.sheets["AssetMgmtRenoStatusReport"].range('A5').end('down').row

    #stores values in column A to X
    #enter new sheet name
    my_values = UA.sheets["AssetMgmtRenoStatusReport"].range("A5:X" + str(Renolen)).options(ndim=2).value

    UA.sheets["Reno"].range('A2:X55000').value = my_values

    sheet = UA.sheets['AssetMgmtRenoStatusReport']
    sheet.delete()

    formula = UA.sheets["Reno"].range("Y2:AC2").options(ndim=2).formula

    #copies column A to C
    UA.sheets["Reno"].range("Y2:AC" + str(Renolen)).formula = formula

    Suitereno.close()

    sheet = UA.sheets['Summary']

    #
    # We have now copied the data and dragged our formulas. The next step is to add aqusitions and get rid of dispos
    #

    accquisition = input("Please enter Yes when new accquisitons or dispositions are added/removed ")

    while (accquisition != "Yes"):
        accquisition = input("You have made an invalid entry, please enter Yes when new accquisitons or dispositions are added/removed ")
    #
    # Now we relink and refresh
    #

    #opens last weeks UA
    old_UA_path = "U:/CDNMF/Reports/Unit Availability/2022/" + formated_old_file_name
    old_UA = xw.Book(old_UA_path, update_links=False)

    #opens new t5w UA
    t5W_new_date = today_date - dt.timedelta(35)
    t5W_new = dt.date.strftime(t5W_new_date,'%#m-%d-%y')
    t5W_path = "U:/CDNMF/Reports/Unit Availability/2022/" + "Unit Availability " + t5W_new + ".xlsm"
    t5W = xw.Book(t5W_path, update_links=False)

    #opens this weeks New Leases
    formatted_new_leases_name = "New Leases " + download_date + ".xlsm"
    new_leases_name = "U:/CDNMF/Reports/New Leases/" + download_date + "/" + formatted_new_leases_name
    new_leases = xw.Book(new_leases_name, update_links=False)

    #opens this weeks Rent Analysis
    formatted_rent_analysis_name = "Rent Analysis " + download_date + ".xlsx"
    rent_analysis_name = "U:/CDNMF/Reports/Rent Analysis Automation/Reports/Rent Analysis/2022/" + formatted_rent_analysis_name
    rent_analysis = xw.Book(rent_analysis_name, update_links=False)

    #opens master list
    master_list_name = "U:/Leasing Data/Master List - CDN MF/" + "00-Master List.xlsm"
    master_list = xw.Book(master_list_name)


    two_weeks_UA_date = today_date - dt.timedelta(14)
    two_weeks_UA = dt.date.strftime(two_weeks_UA_date,'%#m-%d-%y')

    t5W_old_date = today_date - dt.timedelta(42)
    t5W_old = dt.date.strftime(t5W_old_date,'%#m-%d-%y')

    lastweek_rent_newleases_date = today_date - dt.timedelta(7)
    lastweek_rent_newleases = dt.date.strftime(lastweek_rent_newleases_date,'%#m-%#d-%y')

    for sheet in UA.sheets:

        #changes dates for rent analysis and new leases
        UA.sheets[sheet].range("A1:CC5500").api.Replace(lastweek_rent_newleases, download_date)
        #links last weeks UA
        UA.sheets[sheet].range("A1:CC5500").api.Replace(two_weeks_UA, old_file_name)
        #links cell AS13 report
        UA.sheets[sheet].range("A1:CC5500").api.Replace(t5W_old,t5W_new)

    print("\nRelinking finished\n")

    #refreshes with all links open
    UA.api.RefreshAll()

    #closes all extra files
    old_UA.close()
    t5W.close()
    new_leases.close()
    rent_analysis.close()

    #saves our changes
    UA.save()
    master_list.close()
    UA.close()

    bad_dest_dir = "U:/CDNMF/Reports/Unit Availability/2022/" + "unit availability " + new_file_name + ".xlsm"

    #fixes naming convention
    os.rename(bad_dest_dir,dest_dir)

    historical = xw.Book("U:/CDNMF/Reports/Unit Availability/Unit Availability Historical Data.xlsx",update_links = False)

    historical.api.RefreshAll()

    historical.close()

    print("Please see you new UA report at U:/CDNMF/Reports/Unit Availability/2022")

elif (day_of_week == "Tuesday"):
    
    #changes working dir to U drive
    os.chdir("U:/")

    # gets the current working directory (U drive)
    src_dir = os.getcwd()

    # defining the destination directory
    #/CDNMF/Reports/Unit Availability/2022
    dest_dir = src_dir + "/CDNMF/Reports/Unit Availability/2022"

    #now listdir will listen to dest_dir
    os.chdir(dest_dir)

    #adj Tuesday date to monday
    today_date = dt.date.today()  + dt.timedelta(on_cycle) - dt.timedelta(1)

    old_file_name = dt.date.strftime(today_date,'%#m-%d-%y')
    #formats old file name
    formated_old_file_name_bad = "Unit Availability " + old_file_name + ".xlsm"
    formated_old_file_name = "Unit Availability " + old_file_name + "-bad.xlsm"
    
    #renames the file
    os.rename(formated_old_file_name_bad,formated_old_file_name)

    formated_new_file_name = "Unit Availability " + old_file_name + ".xlsm"

    #changes file path to give copied file new name
    dest_dir = dest_dir + "/" + formated_new_file_name

    #copys file under our new name
    shutil.copy(formated_old_file_name, dest_dir)

    #essentially we took Mondays file renamed it with "-bad" at the end, then copied the file and renamed the new file without the "-bad" extension
    #now we remove mondays file
    
    os.remove(formated_old_file_name)

    #opens this weeks UA
    UA = xw.Book(formated_new_file_name, update_links=False)
    
    #goes to availability sheet
    sheet = UA.sheets["Availability"]

    sheet.api.AutoFilterMode = False

    #clears column A to M for availability
    sheet.range("A2:M8000").clear_contents()

    download_date = dt.date.strftime(today_date,'%#m-%#d-%y')

    #adjusts file name
    UA_download = "U:/CDNMF/Unit Availability Download/" + download_date + ".xls"

    #opens UA download
    download = xw.Book(UA_download)

    download.sheets[0].api.AutoFilterMode = False

    #copys sheet to UA report
    download.sheets[0].copy(after=UA.sheets[0])


    UAlen = UA.sheets["report1"].range('A7').end('down').row

    ranger = "A7:C" + str(UAlen)

    #stores values in column A to C
    my_values = UA.sheets["report1"].range(ranger).options(ndim=2).value

    #copies column A to C
    UA.sheets["Availability"].range('A2:C8000').value = my_values

    ranger2 = "F7:O" + str(UAlen)

    #stores values in column F to O
    my_values = UA.sheets["report1"].range(ranger2).options(ndim=2).value

    #copies column F to O
    UA.sheets["Availability"].range('D2:M8000').value = my_values

    sheet = UA.sheets['report1']
    sheet.delete()

    download.close()

    #
    # Now we refresh
    #
    
    last_week_file_date = today_date - dt.timedelta(7)
    last_week_file_name = dt.date.strftime(last_week_file_date,'%#m-%d-%y')

    #opens last weeks UA
    old_UA_path = "U:/CDNMF/Reports/Unit Availability/2022/" + "Unit Availability " + last_week_file_name + ".xlsm"
    old_UA = xw.Book(old_UA_path, update_links=False)

    #opens t5W
    t5W_date = today_date - dt.timedelta(35)
    t5W_name = dt.date.strftime(t5W_date,'%#m-%d-%y')
    AS13_UA_name = "U:/CDNMF/Reports/Unit Availability/2022/" + "Unit Availability " + t5W_name + ".xlsm"
    AS13_UA = xw.Book(AS13_UA_name, update_links=False)

    #opens this weeks New Leases
    formatted_new_leases_name = "New Leases " + download_date + ".xlsm"
    new_leases_name = "U:/CDNMF/Reports/New Leases/" + download_date + "/" + formatted_new_leases_name
    new_leases = xw.Book(new_leases_name, update_links=False)

    #opens this weeks Rent Analysis
    formatted_rent_analysis_name = "Rent Analysis " + download_date + ".xlsx"
    rent_analysis_name = "U:/CDNMF/Reports/Rent Analysis Automation/Reports/Rent Analysis/2022/" + formatted_rent_analysis_name
    rent_analysis = xw.Book(rent_analysis_name, update_links=False)

    #opens master list
    master_list_name = "U:/PSP/IMH/Leasing Data/Master List - CDN MF/" + "00-Master List.xlsm"
    master_list = xw.Book(master_list_name)

    #refreshes with all links open
    UA.api.RefreshAll()

    #closes all extra files
    old_UA.close()
    AS13_UA.close()
    new_leases.close()
    rent_analysis.close()

    # all complete
    print("UA refresh complete")

    #saves our changes
    UA.save()
    master_list.close()
    UA.close()

    bad_dest_dir = "U:/CDNMF/Reports/Unit Availability/2022/" + "unit availability " + old_file_name + ".xlsm"

    #fixes naming convention
    os.rename(bad_dest_dir,dest_dir)

    historical = xw.Book("U:/CDNMF/Reports/Unit Availability/Unit Availability Historical Data.xlsx",update_links = False)

    historical.api.RefreshAll()

    historical.close()

    print("Historical refresh complete")

    print("All processes complete")
    
else:
    print("You have entered an invalid day")
