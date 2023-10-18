
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

today_date = dt.date.today()  + dt.timedelta(on_cycle) - dt.timedelta(2)

#changes working dir to U drive
os.chdir("U:/")

# gets the current working directory (U drive)
src_dir = os.getcwd()

# defining the destination directory
# change to /CDNMF/Reports/Vacancy Summary
dest_dir = src_dir + "/CDNMF/Reports/Vacancy Summary"

#now listdir will listen to dest_dir
os.chdir(dest_dir)

last_week_date = today_date - dt.timedelta(7)
old_file_name = dt.date.strftime(last_week_date,'%#m-%#d-%y')
#formats old file name
formated_old_file_name = "Vacancy Summary-" + old_file_name + ".xlsx"


new_file_name = dt.date.strftime(today_date,'%#m-%#d-%y')
#fomats new file name
formated_new_file_name = "Vacancy Summary-" + new_file_name + ".xlsx"


#changes file path to give copied file new name
dest_dir = dest_dir + "/" + formated_new_file_name

#copys file under our new name
shutil.copy(formated_old_file_name, dest_dir)

#opens historical file
historical = xw.Book("U://CDNMF/Reports/Unit Availability/Unit Availability Historical Data.xlsx",update_links=False)

#opens this weeks Vacancy Summary
VS = xw.Book(formated_new_file_name, update_links=False)

#goes to data sheet
sheet = VS.sheets['Data']

#changes date in cell B1
sheet.range("B1").value = today_date

VS.save()

historical = xw.Book("U:/CDNMF/Reports/Unit Availability/Unit Availability Historical Data.xlsx", update_links=False)

sheet = historical.sheets['UA Historical']

sheet.range("A1:T55000").api.Replace(old_file_name,new_file_name)

VS.close()

bad_dest_dir = "U:/CDNMF/Reports/Unit Availability/2022/" + "vacancy summary-" + old_file_name + ".xlsx"

#fixes naming convention
os.rename(bad_dest_dir,dest_dir)

VS = xw.Book(formated_new_file_name, update_links=False)

# all complete
print("Please see /CDNMF/Reports/Vacancy Summary for your new report")


