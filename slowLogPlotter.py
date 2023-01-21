import datetime
import traceback
import itertools
from tkinter.ttk import Progressbar

import matplotlib.pyplot as plt
import numpy as np
from tkinter import *
from tkcalendar import Calendar
import csv

querytime = []
timestamp = []
time_averages = []
date_array = []
query_name = ""
query_count = []
time_range = 0

def launch_gui():
    def grad_date():
        start_date = cal1.get_date().split('/')
        end_date = cal2.get_date().split('/')
        global query_name
        query_name = str(query_name_tem.get())
        global time_range
        time_range = time_range_var.get()

        start_date = datetime.date(int(start_date[2]), int(start_date[0]), int(start_date[1]))
        end_date = datetime.date(int(end_date[2]), int(end_date[0]), int(end_date[1]))
        date_diff = (end_date - start_date).days

        progress = Progressbar(root, orient=HORIZONTAL,
                               length=150, mode='determinate')

        start_date_formatted = start_date.strftime("%d%m%Y")
        end_date_formatted = end_date.strftime("%d%m%Y")
        start_date_formatted = str(start_date_formatted[0:4]) + str(start_date_formatted[-2:])
        end_date_formatted = str(end_date_formatted[0:4]) + str(end_date_formatted[-2:])
        print(start_date_formatted)
        print(end_date_formatted)

        date1.config(text="Selected Start Date is: " + start_date_formatted)
        date2.config(text="Selected end Date is: " + end_date_formatted)
        progress.grid(row=6,columnspan=2)
        root.update()
        filearray = generate_filename_array(start_date, end_date)
        build_time_array(filearray, progress, root, date_diff)

        #for x in filearray:
        #    print(x)
        root.destroy()
    root = Tk()

    # Set geometry
    # root.geometry("600x600")

    query_label = Label(root, text="Select Query to run analysis", width="40", padx="40", font=("Nunito Sans", 10))
    global query_name
    query_name_tem = StringVar()
    query_name_tem.set("Select")
    query_list = [
        "ZipHealth.get_provider_list_for_a_metric",
        "ZipHealth.calculate_care_gap_list_for_org",
        "ZipHealth.my_patient_into_temporary_table",
        "ZipHealth.unified_internal_external_provider_search",
        "ZipHealth.unified_global_search",
        "ZipHealth.get_patient_events",
        "ZipHealth.get_contact_log",
        "ZipHealth.sp_patient_list",
        "ZipHealth.get_people_list",
        "ZipHealth.get_hl7_lab_details",
        "ZipHealth.get_hospital_activity_list",
        "ZipHealth.get_user_list",
        "ZipHealth.get_ehr_problem_list",
        "ZipHealth.handle_hospital_activity",
        "ZipHealth.get_encounter_details_for_provider",
        "ZipHealth.get_csm_task_list",
        "ZipHealth.smart_search_ehr_problem",
        "ZipHealth.get_appointment_schedule",
        "ZipHealth.get_calendar_events",
        "ZipHealth.get_csm_report",
        "ZipHealth.get_superbill_list",
        "ZipHealth.populate_coding_tool_chart_list",
        "ZipHealth.get_quality_coding_sheet",
        "ZipHealth.calculate_provider_metric_score",
        "ZipHealth.get_csm_registries",
        "ZipHealth.ccd_bulk_review",
        "ZipHealth.ccda_member_matching_event",
        "ZipHealth.computation_independent_event",
        "ZipHealth.smart_search_ehr_problem",
        "new_ui_ehr_patient_add_encounter_parent",
        "ZipHealth.ehr_processing_event",
        "ZipHealth.get_organisation_list_for_drop_down",
        "ZipHealth.get_API_details",
        "ZipHealth.unified_internal_external_provider_search",
        "ZipHealth.user_email_notification",
        "ZipHealth._care_opp_details"

    ]
    query_drop = OptionMenu(root, query_name_tem, *query_list)
    to_date = datetime.date.today() + datetime.timedelta(days=1)
    from_date = to_date + datetime.timedelta(days=-60)

    cal1 = Calendar(root, selectmode='day', year=int(from_date.year), month=int(from_date.month), day=int(from_date.day))
    cal2 = Calendar(root, selectmode='day', year=int(to_date.year), month=int(to_date.month), day=int(to_date.day))

    query_label.grid(row=0, columnspan=2)
    query_drop.grid(row=1, columnspan=2)
    cal1.grid(row=2, column=0)
    cal2.grid(row=2, column=1)

    Button(root, text="Analyze Data",
           command=grad_date).grid(row=4, columnspan=2)
    time_range_var = IntVar()
    all_day_radio = Radiobutton(root, text="All Day", variable=time_range_var, value=0, font=("Nunito Sans", 10))
    business_hours_radio = Radiobutton(root, text="Business Hours only", variable=time_range_var, value=1, font=("Nunito Sans", 10))

    all_day_radio.grid(row=3, column=0)
    business_hours_radio.grid(row=3, column=1)


    date1 = Label(root, text="")
    date2 = Label(root, text="")
    date1.grid(row=5, column=0)
    date2.grid(row=5, column=1)



    # Execute Tkinter
    root.mainloop()



def generate_filename_array(start_date, end_date):
    working_date = start_date
    global date_array
    filename_array = []
    flag=True
    while flag:
        working_date_formatted = working_date.strftime("%d%m%Y")
        working_date_formatted = str(working_date_formatted[0:4])+str(working_date_formatted[-2:])
        filename_array.append("prod_slow_logs_"+working_date_formatted+".csv")
        date_array.append(working_date_formatted)
        working_date+=datetime.timedelta(days=1)

        if working_date == end_date:
            flag = False
    return filename_array


def build_time_array(filearray, progress, root, date_diff):
    global querytime
    global timestamp
    global time_averages
    global date_array
    global query_name
    global query_count

    for filear in filearray:
        print(filear)
    for dates in date_array:
        print(dates)
    progressval = 1
    for name in filearray:
        try:
            file = open('C:\\Slow_log_data\\'+name)
            type(file)
            csvreader = csv.reader(file)
            rows = []
            print("Found file:"+name)
            for row in csvreader:
                #if row[0] == query_name:
                if query_name in row[0]:
                    querytime.append(float(row[1]))
                    #print(float(row[1]))
                    timestamp.append(row[2])
            print("File "+name+" Loaded Successfully")
            print("Query time count:"+str(len(querytime)))
            print("Timestamp count:"+str(len(timestamp)))
            progress['value'] = (progressval/date_diff)*100
            root.update_idletasks()
            time_average = 0
            business_hour_count = 0
            progressval += 1
            for time_taken, current_time in zip(querytime, timestamp):
                if time_range == 1:
                    hour = int(current_time.split("T")[1].split(":")[0])
                    if 7 < hour < 18:
                        #print("Hour "+str(hour))
                        time_average = time_average + time_taken
                        #print("Business hours")
                        business_hour_count+=1
                        continue
                    else:
                        continue
                time_average = time_average + time_taken
                #print("All day")
            #print(time_average)
            #time_averages.append(time_average)
            if time_range == 1:
                time_averages.append(time_average/float(business_hour_count))
                query_count.append(business_hour_count)
            else:
                time_averages.append(time_average/float(len(querytime)))
                query_count.append(len(querytime))
            querytime.clear()
            timestamp.clear()



        except Exception as e:
            print(e)
            traceback.print_exc()
            print("NEED TO REMOVE:"+name[-10:-4])
            date_array.remove(name[-10:-4])
            querytime.clear()
            timestamp.clear()
            continue






launch_gui()
count =0
for i in querytime:
    if i>30:
        count+=1
print(count)

for i,j,k in zip(date_array, time_averages, query_count):
    print(str(i)+" : "+str(j)+" : "+str(k))


x=np.array(date_array)
y1=np.array(time_averages)
y2=np.array(query_count)

fig, ax1 = plt.subplots()
ax1.set_xlabel('DATES', color='red')
ax1.set_ylabel('QUERY TIME AVERAGES - '+query_name, color='red')
ax1.tick_params(axis ='y', labelcolor = 'red')
ax1.plot(x,y1, color="red")
ax1.tick_params(axis='x',rotation=45)

plt.grid()
ax2 = ax1.twinx()
ax2.set_ylabel('QUERY COUNT - '+query_name, color="blue")
ax2.plot(x,y2, color="blue")
ax2.tick_params(axis ='y', labelcolor = 'blue')

plt.show()

