
import os
import shutil
import xlwings as xw
import datetime as dt

#this input helps us track where we are in the week
on_cycle = input("If you are on-cycle enter 0, if not please enter the number of days missed since Monday ")

#this loop ensures a value is entered
while (on_cycle == ""):
    on_cycle = input("You have not entered a value, please try again ")

#we make this a numebr so we can use it later
on_cycle = int(on_cycle) * -1

os.chdir("U:/python tester folder")

#opens the UA download saved in the python tester folder
formated_download_file_name = "U:/python tester folder/" + "MM download.xlsx"

#opens this weeks download
download = xw.Book(formated_download_file_name, update_links=False)

#goes to main sheet
sheet = download.sheets['Report1']

#gets the number of rows in the download
downloadlen = sheet.range('A7').end('down').row

#stores values in column B
my_values = sheet.range('B7:B'+str(downloadlen)).options(ndim=2,numbers=int).value

#repastes them as ints.
download.sheets["Report1"].range('B7:B'+str(downloadlen)).value = my_values

#Uses column D to create property + unit type Ids
sheet.range("D7").value = "=A7&B7"

#copys formula
formula = download.sheets["report1"].range("D7").options(ndim=2).formula

#drags formula down in column D
download.sheets["report1"].range("D7:D" + str(downloadlen)).formula = formula

#we get todays date, backdate to monday, and then add the on cycle to account for short weeks
DD_folder_date = dt.date.today() - dt.timedelta(3) + dt.timedelta(on_cycle)

#fixes the format of the date to match DD naming convention
td_file_date_DD = dt.date.strftime(DD_folder_date,'%m-%d-%y')

DD_file_name = "DD Weekly Update " + td_file_date_DD + ".xlsx"

#puts together DD path
DD_name = "U:/Leasing Data/DD - Semi-Monthly Update/" + str(DD_folder_date) + "/" + DD_file_name
#opens DD
DD = xw.Book(DD_name,update_links= False)

DD_sheet = DD.sheets['Rent Roll']
#gets amount of entries in DD rent roll
DDlen = DD_sheet.range('A2').end('down').row

#inserts a column 
DD_sheet.range('A:A').insert()

#creates a property + unit type unique ID
DD_sheet.range("A2").value = "=B2&C2"

formula = DD_sheet.range("A2").options(ndim=2).formula

#drags ID formula down
DD_sheet.range("A2:A" + str(DDlen)).formula = formula

#lookup array from DD
lookup_array = "\'[" + DD_file_name + "]Rent Roll\'!" + "$A$2:$A$" + str(DDlen)
#return array in DD, we are just returning prop code
return_array = "\'[" + DD_file_name + "]Rent Roll\'!" + "$B$2:$B$" + str(DDlen)

#sheet refers to the UA download, essentially we are entering the xlookup formula in column E of the download to index with DD
sheet.range('E7').value = "=XLOOKUP(D7," + lookup_array +"," + return_array + ")"

formula = sheet.range("E7").options(ndim=2).formula
#drags xlookup formula down
sheet.range("E7:E" + str(downloadlen)).formula = formula

del_list=[]
downlen = downloadlen-6

for i in range(downlen):
    if sheet[i+6,4].value == None:
        del_list.append(i+6)
    else:
        continue
counter = 0

for n in del_list:
    sheet[n+counter,0:].delete()
    counter = counter - 1

download.save()
DD.close()
#changes working dir to U drive
os.chdir("U:/")

# defining the destination directory
dest_dir = "U:/" + "CDNMF/Reports/Minimum Market Rents"

#now listdir will listen to dest_dir
os.chdir(dest_dir)

today_date = dt.date.today()
last_week_date = dt.date.today() - dt.timedelta(7)

#collects old file name from user
old_folder_name = dt.date.strftime(last_week_date,'%m.%d.%Y')


#collects new file name from user
new_folder_name = dt.date.strftime(today_date,'%m.%d.%Y')

format_old_folder_name = dest_dir + "/" + old_folder_name

#changes file path to give copied file new name
dest_dir = dest_dir + "/" + new_folder_name

#copys folder under our new name
shutil.copytree(format_old_folder_name, dest_dir)

H_file = ""
H_file = "Company data/" + new_folder_name + "/Minimum Market Rents - " + old_folder_name + " - Hard Coded.xlsx"
dull_file = dest_dir + "~$Minimum Market Rents - 03.15.2021.xlsx"

# removes last years files
os.remove(H_file)

try:
    os.remove(dull_file)
    
except FileNotFoundError:
    print("Dull file not here")
    pass

old_file_name = "Minimum Market Rents - " + old_folder_name + ".xlsx"
new_file_name = "Minimum Market Rents - " + new_folder_name + ".xlsx"

os.chdir(dest_dir)

os.rename(old_file_name,new_file_name)

MM = xw.Book(new_file_name, update_links=False)

sheet = MM.sheets['Summary']

#updates last report date in summary tab
sheet.range('s2').value = (dt.date.today() - dt.timedelta(7))

two_weeks_MM_bad = dt.date.today() - dt.timedelta(14)
two_weeks_MM = dt.date.strftime(two_weeks_MM_bad,'%m.%d.%Y')

sheet = MM.sheets['Unit Availability']
sheet.api.AutoFilterMode = False

#clears column A to O for Minimum market rents availability
sheet.range("A2:O8000").clear_contents()
sheet.range("P3:S8000").clear_contents()

#copys sheet to MM rents report
download.sheets[0].copy(after=MM.sheets[0])

#stores values in column A to C
my_values = MM.sheets["report1"].range("A7:O" + str(downloadlen)).options(ndim=2).value

#copies column A to C
sheet.range("A2:O" + str(downloadlen)).value = my_values

MM.sheets['report1'].delete()

formula = sheet.range("P2:S2").options(ndim=2).formula

#copies column A to C
MM.sheets["Unit Availability"].range("P2:S" + str(downloadlen)).formula = formula

input("Please enter Yes when you have addedd/removed accquisitons/dispos ")

old_MM_name = "U:/CDNMF/Reports/Minimum Market Rents/" + old_folder_name + "/" + old_file_name

old_MM = xw.Book(old_MM_name, update_links=False)

master_list_name = "U:/Leasing Data/Master List - CDN MF/" + "00-Master List.xlsm"
master_list = xw.Book(master_list_name)

#re links MM from last week
MM.sheets['Summary'].range("T3:AH471").api.Replace(two_weeks_MM, old_folder_name)

#refreshes with all links open
MM.api.RefreshAll()

old_MM.close()
master_list.close()
download.close()

MM.save()

new_file_name_HC = "Minimum Market Rents - " + new_folder_name + " - Hard Coded.xlsx"

shutil.copy(new_file_name,new_file_name_HC)

MM.close()

MM_HC = xw.Book(new_file_name_HC, update_links=False)

MM_HC.sheets['Summary'].range("A1:AI550").copy()

MM_HC.sheets['Summary'].range("A1:AI550").paste(paste="values")

MM_HC.sheets['Unit Availability'].delete()

MM_HC.sheets['Unit Type Mapping'].delete()

MM_HC.save()

MM_HC.close()

print("Hard coded version complete")

bad_file_name = "minimum market rents - " + new_folder_name + ".xlsx"
os.rename(bad_file_name,new_file_name)

bad_file_name_HC = "minimum market rents - " + new_folder_name + " - hard coded.xlsx"

os.rename(bad_file_name_HC, new_file_name_HC)

print("All processes complete")






