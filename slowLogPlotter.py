import datetime
import traceback
import itertools
from tkinter.ttk import Progressbar
#import babel

from babel import numbers

import matplotlib.pyplot as plt
import numpy as np
from tkinter import *

from matplotlib import ticker
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
        custom_query_name = query_name_custom.get().strip()
        if custom_query_name != "Enter Custom Query":
            query_name = custom_query_name
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
        progress.grid(row=7,columnspan=2)
        root.update()
        filearray = generate_filename_array(start_date, end_date)
        build_time_array(filearray, progress, root, date_diff)
        generate_trendgraph()

        #for x in filearray:
        #    print(x)
        root.destroy()

    def generate_histogram():
        start_date = cal1.get_date().split('/')
        end_date = cal2.get_date().split('/')
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
        #print(start_date_formatted)
        #print(end_date_formatted)

        date1.config(text="Selected Start Date is: " + start_date_formatted)
        date2.config(text="Selected end Date is: " + end_date_formatted)

        progress.grid(row=7, columnspan=2)
        root.update()

        query_based_processed_list = []
        for histo_query_name in query_list_histo:
            print(str(int(query_list_histo.index(histo_query_name)+1))+".processing proc: "+histo_query_name)


            filearray = generate_filename_array(start_date, end_date)
            build_time_array_histogram(histo_query_name, filearray, progress, root, date_diff)
            time_total = 0
            query_total = 0

            for j, k in zip(time_averages, query_count):
                time_total = time_total + j
                query_total = query_total + k

            try:
                query_based_processed_list.append([histo_query_name, time_total/float(query_total), query_total])
                print([[histo_query_name, time_total/float(query_total), query_total]])
            except ZeroDivisionError as e:
                query_based_processed_list.append([histo_query_name, 0, query_total])

            time_averages.clear()
            query_count.clear()
            date_array.clear()
        print(query_based_processed_list)
        generate_histogram_graph(query_based_processed_list)




    root = Tk()

    # Set geometry
    # root.geometry("600x600")

    query_label = Label(root, text="Select Query or Enter custom query to run analysis", width="40", padx="40", font=("Nunito Sans", 10))
    global query_name
    query_name_tem = StringVar()
    query_name_tem.set("Select")
    query_or_label = Label(root, text="or", width="40", padx="40", font=("Nunito Sans", 10))
    query_name_custom = Entry(root, width=20)
    query_list_histo = [
        "ZipHealth.automatic_extract_controll",
        "ZipHealth.HP_patient_worker_event_composite_2",
        "ZipHealth.HP_patient_worker_event_2",
        "ZipHealth.worker_event_care_ops_count_for_refresh",
        "ZipHealth.support_worker_event_rel_18",
        "ZipHealth.support_worker_event_4",
        "ZipHealth.HP_patient_worker_event_composite_1",
        "ZipHealth.support_worker_event_2",
        "ZipHealth.provider_worker_event_1",
        "ZipHealth.support_worker_event",
        "ZipHealth.support_worker_event_5",
        "ZipHealth.support_worker_event_3",
        "ZipHealth.support_worker_event_rel_10",
        "ZipHealth.computation_independent_event",
        "ZipHealth.HP_patient_worker_event_1",
        "ZipHealth.support_worker_event_rel_11",
        "ZipHealth.support_worker_event_rel_14",
        "ZipHealth.support_worker_event_rel_12",
        "ZipHealth.support_worker_event_rel_17",
        "ZipHealth.support_worker_event_rel_15",
        "ZipHealth.support_worker_event_rel_16",
        "ZipHealth.computation_controller_event",
        "ZipHealth.HP_patient_worker_event_3",
        "ZipHealth.support_worker_event_rel_13",
        "ZipHealth.support_worker_event_rel_19",
        "ZipHealth.ccda_member_matching_event",
        "ZipHealth.provider_worker_event_2",
        "Analytics.insert_risk_score_level_info_sweep",
        "ZipHealth.daily_refresh_worker",
        "ZipHealth.provider_metric_cache_bulk_compute_event",
        "ZipHealth.patient_worker_event_hcc_hnet",
        "Analytics.refresh_redshift_aca_hcc_continuation",
        "ZipHealth.support_worker_event_6",
        "Analytics.populate_patient_covered_days",
        "Analytics.refresh_redshift_tables",
        "Analytics.process_redshift_cost_categorization",
        "Analytics.refresh_cost_trends",
        "ZipHealth.ccd_bulk_review",
        "ZipHealth.measure_extract_in_details_parallel_run",
        "Analytics.refresh_patient_hierarchy",
        "Analytics.refresh_patient_detail_mapping",
        "Analytics.refresh_quality_measures",
        "Analytics.populate_hpmg_specialty_bonus",
        "ZipHealth.search_tag_add_update_proc",
        "ZipHealth.cache_purge_for_po",
        "Analytics.refresh_patient_quality",
        "Analytics.populate_hcc_dx_reconciliation",
        "ZipHealth.data_processing_with_unique_id",
        "Analytics.refresh_patient_org_hierarchy",
        "ZipHealth.populate_weekly_performance_data",
        "ZipHealth.migrate_FDB_data_changes_to_ZipHealth_tables",
        "ZipHealth.populate_coding_tool_chart_list",
        "ZipHealth.ccd_clearance",
        "Analytics.populate_risk_other_info",
        "Analytics.populate_cohort_builder",
        "ZipHealth.get_customer_weekly_diff",
        "ZipHealth.patient_worker_event_composite",
        "ZipHealth.create_computation_weekly_metadata_prod",
        "Analytics.populate_notable_events_flag",
        "Analytics.process_hospital_risk_data",
        "Analytics.create_refresh_log_delta",
        "ZipHealth.project_provider_performance_data",
        "ZipHealth.populate_cf_rel_tables_from_ccd_data",
        "ZipHealth.patient_worker_utilization",
        "ZipHealth.patient_worker_event_4",
        "ZipHealth.get_user_list",
        "ZipHealth.refresh_patient_metric_cache",
        "ZipHealth.refresh_global_search_table_splits",
        "ZipHealth.HP_refresh_patient_metric_cache",
        "ZipHealth.refresh_patient_utilization_cache",
        "ZipHealth.patient_worker_event_composite_6",
        "ZipHealth.patient_worker_event_5",
        "ZipHealth.patient_worker_utilization_2",
        "ZipHealth.provider_metric_cache_bulk_compute",
        "ZipHealth.refresh_patient_metric_cache_composite",
        "ZipHealth.get_API_details",
        "ZipHealth.patient_worker_utilization_3",
        "ZipHealth.patient_worker_event_composite_2",
        "ZipHealth.patient_worker_event_composite_4",
        "ZipHealth.HP_refresh_patient_metric_cache_composite",
        "ZipHealth.patient_worker_event_composite_3",
        "ZipHealth.patient_worker_event_composite_5",
        "ZipHealth.patient_worker_event_count_composite_event",
        "ZipHealth.refresh_patient_metric_cache_event_based_composite",
        "ZipHealth.patient_worker_utilization_4",
        "ZipHealth._get_metric_score_for_providers_rel",
        "Analytics.refresh_redshift_suspect_tables_aca",
        "ZipHealth.generate_bridge_MM_dump_for_PO_test",
        "ZipHealth.patient_worker_event_count_composite_event_2",
        "ZipHealth.populate_other_chart_suggestions_problems",
        "Analytics.populate_final_patient_next_appointment_table",
        "ZipHealth.patient_worker_event_count_composite_event_3",
        "ZipHealth.patient_worker_attribute_2",
        "ZipHealth.compute_eng_measure_cache",
        "ZipHealth.my_patient_into_temporary_table",
        "ZipHealth.patient_worker_event_quick",
        "ZipHealth.patient_worker_event",
        "ZipHealth.HP_patient_worker_event_composite_3",
        "ZipHealth.patient_worker_event_count_composite_event_4",
        "ZipHealth.patient_worker_event_count_composite_event_6",
        "ZipHealth.patient_worker_event_record_based_2",
        "ZipHealth.patient_worker_event_hcc",
        "ZipHealth.automatic_job_schedule",
        "ZipHealth.refresh_patient_metric_cache_record_based",
        "ZipHealth.load_calinx_pharmacy_for_customer",
        "ZipHealth.patient_worker_event_3",
        "ZipHealth.patient_worker_event_record_based_3",
        "ZipHealth.patient_worker_event_count_composite_event_5",
        "ZipHealth.patient_worker_event_record_based_1",
        "ZipHealth.refresh_patient_attributes_cache_2",
        "ZipHealth.computation_controller",
        "ZipHealth.calculate_provider_metric_score",
        "ZipHealth.independent_EHR_feed_computation",
        "ZipHealth.smart_search_ehr_problem_template",
        "ZipHealth.unified_internal_external_provider_search",
        "ZipHealth.get_login_params_for_user",
        "ZipHealth.get_patient_list_for_member_support",
        "ZipHealth.populate_feature_for_member_engagement_rolling",
        "ZipHealth.patient_worker_attribute",
        "Analytics.refresh_all_user_details",
        "ZipHealth.refresh_patient_attributes_cache",
        "Analytics.refresh_patient_detail_mapping_sdoh",
        "Analytics.populate_member_enrollment_count",
        "ZipHealth.generate_export_event",
        "ZipHealth.patient_worker_event_hcc_p4",
        "Analytics.refresh_patient_detail_mapping_address_info",
        "ZipHealth.generate_member_match_dump_for_bridge",
        "ZipHealth._get_metric_score_for_providers",
        "Analytics.populate_active_member_count_for_billing",
        "ZipHealth.chart_chase_error_processing",
        "Analytics.populate_hcc_list_yearly",
        "ZipHealth.populate_provider_network_comparison",
        "Analytics.find_quality_attribution",
        "Analytics.refresh_patient_detail_mapping_phone_email",
        "ZipHealth.calculate_provider_metric_score_bulk_compute",
        "Analytics.refresh_all_pcp_map_quality",
        "Analytics.refresh_user_performances",
        "Analytics.refresh_provider_quality",
        "ZipHealth.manage_chart_list_export",
        "Analytics.refresh_patient_county",
        "Analytics.refresh_provider_cache",
        "ZipHealth.get_csg_di_details",
        "Analytics.process_populate_cms_file_dates",
        "ZipHealth.ccda_member_and_provider_matching",
        "ZipHealth.cpt_computation_suspect_supp_od",
        "ZipHealth.calculate_care_opps_count",
        "Analytics.refresh_org_info",
        "ZipHealth.refresh_provider_cache_bulk_compute",
        "ZipHealth.get_encounter_details_for_provider",
        "ZipHealth.unified_global_search",
        "Analytics.refresh_org_hierarchy",
        "Analytics.track_hcc_visit_yearly",
        "ZipHealth.manage_chart_list_export_data",
        "FDB.LoadMappingFile",
        "FDB.LoadMappingTable",
        "ZipHealth.writeback_provider_id_in_ccd_tables",
        "ZipHealth.calculate_care_gap_list_for_org",
        "ZipHealth.refresh_stale_patient_hcc_cache",
        "ZipHealth.extract_bridge_data",
        "ZipHealth.process_ccd_data_for_claims",
        "ZipHealth._care_opp_details",
        "ZipHealth.icdpcs_computation_suspect_supp_od",
        "ZipHealth.get_patient_list",
        "ZipHealth.preprocess_calinx_raw_table",
        "ZipHealth.get_csm_patient_specific_task_list",
        "Analytics.refresh_customer_outreach_data",
        "ZipHealth.match_member_for_customer_calinx",
        "Analytics.insert_hcc_history_monthly",
        "ZipHealth.get_encounter_location",
        "ZipHealth.rx_computation_suspect_ehr_od",
        "ZipHealth.refresh_stale_patient_hcc_cache_hnet",
        "ZipHealth.cache_purge_for_hp",
        "ZipHealth.dx_computation_suspect_supp_od",
        "ZipHealth.get_contact_log",
        "ZipHealth.ehr_based_incremental_attribute_computation_doc",
        "ZipHealth.refresh_support_cache_for_stale_results_rel",
        "ZipHealth.refresh_support_cache_for_stale_results",
        "ZipHealth.measure_related_excluded_patient_list",
        "ZipHealth.update_active_inactive_severity",
        "ZipHealth.get_provider_list_for_a_metric",
        "ZipHealth.refresh_care_ops_count",
        "ZipHealth.patient_details",
        "ZipHealth.get_ehr_vaccine_info",
        "ZipHealth.smart_search_ehr_problem",
        "ZipHealth.get_load_eligible_rows_for_RX",
        "ZipHealth.refresh_provider_cache_for_stale_results_others",
        "Analytics.calculate_hpmg_sb_patient_grx",
        "ZipHealth.lb_computation_suspect_supp_od",
        "ZipHealth.handle_hospital_activity",
        "ZipHealth.splty_computation_suspect_supp_od",
        "ZipHealth.preparing_the_export",
        "ZipHealth.generic_logger",
        "Analytics.track_hcc_visit_yearly_initial",
        "ZipHealth.get_csm_registries",
        "ZipHealth.outbound_member_extract_for_all_PO_customer",
        "ZipHealth.process_hospital_data",
        "ZipHealth.ePayment_webhook_process",
        "ZipHealth.get_csm_task_list",
        "ZipHealth.get_people_list",
        "ZipHealth.add_or_edit_csm_form_for_task",
        "Analytics.refresh_analytics_delta",
        "ZipHealth.get_ehr_problem_list",
        "ZipHealth.get_patient_events",
        "ZipHealth.compute_eng_measure",
        "Analytics.delete_amp_extra_data",
        "ZipHealth.get_support_activity",
        "ZipHealth.customer_calinx_pbm_delete",
        "ZipHealth.customer_V2_load_schedules",
        "ZipHealth.get_user_activity_extract",
        "ZipHealth.get_user_outbound_extract_for_customer_v2",
        "ZipHealth.get_risk_coding_sheet",
        "Analytics.refresh_sb_quality_measures",
        "ZipHealth.get_all_noncompliant_patients_for_single_provider_export",
        "ZipHealth.remove_highlights_for_detached_document",
        "ZipHealth.chart_warning_and_suggestions",
        "ZipHealth.update_connect_account_map",
        "ZipHealth.hill_add_patient_facility_claim_information_batch_1",
        "ZipHealth.HCC_lookup",
        "ZipHealth.update_chart_response_task_split",
        "ZipHealth.add_or_edit_ehr_vitals",
        "ZipHealth.get_batch_share_list"
    ]
    query_list_histo = [
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
        "ZipHealth.get_organisation_list_for_drop_down",
        "ZipHealth.get_API_details",
        "ZipHealth.user_email_notification",
        "ZipHealth.patient_details",
        "ZipHealth.get_banner",
        "ZipHealth.calculate_care_gap_list_for_org_hcc",
        "ZipHealth.manage_chart_list_export",
        "ZipHealth._care_opp_details",
        "ZipHealth.set_contact_type_email",
        "ZipHealth.get_support_activity"
    ]

    query_list_old = [
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
        "ZipHealth.smart_search_ehr_problem",
        "new_ui_ehr_patient_add_encounter_parent",
        "ZipHealth.ehr_processing_event",
        "ZipHealth.get_organisation_list_for_drop_down",
        "ZipHealth.get_API_details",
        "ZipHealth.user_email_notification",
        "ZipHealth._care_opp_details",
        "ZipHealth.set_contact_type_email",
        "ZipHealth.get_support_activity",
        "ZipHealth.computation_independent_event",
        "ZipHealth.get_csg_di_details"
    ]
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
        "ZipHealth.get_organisation_list_for_drop_down",
        "ZipHealth.get_API_details",
        "ZipHealth.user_email_notification",
        "ZipHealth.patient_details",
        "ZipHealth.get_banner",
        "ZipHealth.calculate_care_gap_list_for_org_hcc",
        "ZipHealth.manage_chart_list_export",
        "ZipHealth._care_opp_details",
        "ZipHealth.set_contact_type_email",
        "ZipHealth.get_support_activity",
        "ZipHealth.get_csg_di_details",
        "ZipHealth.get_stickets",
        "ZipHealth.get_attempted_gap_closure_list",
        "ZipHealth.get_patient_list",
        "ZipHealth.get_csm_patient_specific_task_list",
        "ZipHealth.get_batch_share_list",
        "ZipHealth.get_patient_list_for_member_support"
    ]
    query_drop = OptionMenu(root, query_name_tem, *query_list)
    to_date = datetime.date.today() + datetime.timedelta(days=1)
    from_date = to_date + datetime.timedelta(days=-60)

    cal1 = Calendar(root, selectmode='day', year=int(from_date.year), month=int(from_date.month), day=int(from_date.day))
    cal2 = Calendar(root, selectmode='day', year=int(to_date.year), month=int(to_date.month), day=int(to_date.day))

    query_label.grid(row=0, columnspan=2)
    query_drop.grid(row=1, columnspan=2)
    query_name_custom.grid(row=2, columnspan=2)
    query_name_custom.insert(0, "Enter Custom Query")
    cal1.grid(row=3, column=0)
    cal2.grid(row=3, column=1)

    Button(root, text="Analyze Data",
           command=grad_date).grid(row=5, column=0)
    Button(root, text="Generate top procs Histogram", command=generate_histogram).grid(row=5, column=1)
    time_range_var = IntVar()
    all_day_radio = Radiobutton(root, text="All Day", variable=time_range_var, value=0, font=("Nunito Sans", 10))
    business_hours_radio = Radiobutton(root, text="Business Hours only", variable=time_range_var, value=1, font=("Nunito Sans", 10))

    all_day_radio.grid(row=4, column=0)
    business_hours_radio.grid(row=4, column=1)


    date1 = Label(root, text="")
    date2 = Label(root, text="")
    date1.grid(row=6, column=0)
    date2.grid(row=6, column=1)



    # Execute Tkinter
    root.title("Slow Log Trend Plotter")
    root.iconbitmap("assets/icon.ico")
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


def build_time_array_histogram(histo_query_name, filearray, progress, root, date_diff):
    global querytime
    global timestamp
    global time_averages
    global date_array
    global query_name
    global query_count

    # for filear in filearray:
    #     print(filear)
    # for dates in date_array:
    #     print(dates)
    progressval = 1
    for name in filearray:
        try:
            file = open('C:\\Slow_log_data\\'+name)
            type(file)
            csvreader = csv.reader(file)
            rows = []
            #print("Found file:"+name)
            for row in csvreader:
                #if row[0] == query_name:
                if histo_query_name == row[0]:
                    querytime.append(float(row[1]))
                    #print(float(row[1]))
                    timestamp.append(row[2])
            #print("File "+name+" Loaded Successfully")
            #print("Query time count:"+str(len(querytime)))
            #print("Timestamp count:"+str(len(timestamp)))
            progress['value'] = (progressval/date_diff)*100
            root.update_idletasks()
            time_average = 0
            business_hour_count = 0
            progressval += 1
            for time_taken, current_time in zip(querytime, timestamp):
                if time_range == 1:
                    hour = int(current_time.split("T")[1].split(":")[0])
                    if 7 < hour < 17:
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
                time_averages.append(time_average)
                query_count.append(business_hour_count)
            else:
                time_averages.append(time_average)
                query_count.append(len(querytime))
            querytime.clear()
            timestamp.clear()



        except Exception as e:
            print(e)
            #traceback.print_exc()
            print("NEED TO REMOVE:"+name[-10:-4])
            date_array.remove(name[-10:-4])
            querytime.clear()
            timestamp.clear()
            continue
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
                if query_name == row[0]:
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
                    if 7 < hour < 17:
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



def generate_histogram_graph(data):
    # Sort data by average runtimes in descending order
    data_sorted = sorted(data, key=lambda x: x[1], reverse=True)

    # Separating the sorted data
    query_names_histo = [item[0] for item in data_sorted]
    average_runtimes = [item[1] for item in data_sorted]
    total_executions = [item[2] for item in data_sorted]

    query_names_histo = query_names_histo[:24]
    average_runtimes = average_runtimes[:24]
    total_executions = total_executions[:24]

    for index, indi_name in enumerate(query_names_histo):
        query_names_histo[index] = indi_name.replace("ZipHealth.", "")

    # Creating the plot
    fig, ax = plt.subplots(figsize=(10, 8))

    # Histogram for Average Runtimes
    bars = ax.bar(query_names_histo, average_runtimes, color='skyblue')
    ax.set_title('Total Query count Over last Month')
    ax.set_ylabel('Average Runtime (s)')
    ax.set_xticks(range(len(query_names_histo)))
    ax.set_xticklabels(query_names_histo, rotation=45, ha='right')

    #ax.yaxis.grid(True)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(base=5))  # Adjust the base value as needed for more or less granularity


    # Adding total_executions on top of the bars
    for bar, total_exe in zip(bars, total_executions):
        height = bar.get_height()
        ax.annotate(f'{total_exe}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom')

    plt.tight_layout()
    plt.show()


def generate_trendgraph():
    dummy_variable = 0
    for i, j, k in zip(date_array, time_averages, query_count):
        print(str(i) + " : " + str(j) + " : " + str(k))

    x = np.array(date_array)
    y1 = np.array(time_averages)
    y2 = np.array(query_count)


    fig, ax1 = plt.subplots()
    ax1.set_xlabel('DATES', color='red')
    ax1.set_ylabel('QUERY TIME AVERAGES - ' + query_name, color='red')
    ax1.tick_params(axis='y', labelcolor='red')
    ax1.plot(x, y1, color="red")
    ax1.tick_params(axis='x', rotation=45)

    plt.grid()
    ax2 = ax1.twinx()
    ax2.set_ylabel('QUERY COUNT - ' + query_name, color="blue")
    ax2.plot(x, y2, color="blue")
    ax2.tick_params(axis='y', labelcolor='blue')

    plt.show()
test_data = [['ZipHealth.get_provider_list_for_a_metric', 5.410423581046125, 15849], ['ZipHealth.calculate_care_gap_list_for_org', 12.785645447244155, 27741], ['ZipHealth.my_patient_into_temporary_table', 6.178776664762744, 9104], ['ZipHealth.unified_internal_external_provider_search', 3.8704676688998103, 29704], ['ZipHealth.unified_global_search', 4.136052327591917, 132952], ['ZipHealth.get_patient_events', 4.581461661501092, 11938], ['ZipHealth.get_contact_log', 8.050035718359318, 14238], ['ZipHealth.sp_patient_list', 6.129984720988269, 4007], ['ZipHealth.get_people_list', 3.8727845324097108, 40204], ['ZipHealth.get_hl7_lab_details', 18.41434351315789, 228], ['ZipHealth.get_hospital_activity_list', 8.53004168271923, 18211], ['ZipHealth.get_user_list', 55.837320792047336, 15215], ['ZipHealth.get_ehr_problem_list', 4.920127369350496, 3341], ['ZipHealth.handle_hospital_activity', 3.7349435505804323, 603], ['ZipHealth.get_encounter_details_for_provider', 23.553172687763716, 6162], ['ZipHealth.get_csm_task_list', 4.612833458923082, 32853], ['ZipHealth.smart_search_ehr_problem', 6.252993811064718, 958], ['ZipHealth.get_appointment_schedule', 7.028981530351436, 313], ['ZipHealth.get_calendar_events', 3.2557707114241, 3195], ['ZipHealth.get_csm_report', 7.598494504833514, 931], ['ZipHealth.get_superbill_list', 12.120239407128258, 7828], ['ZipHealth.populate_coding_tool_chart_list', 6.56150880397753, 77083], ['ZipHealth.get_quality_coding_sheet', 4.2387242185129, 659], ['ZipHealth.calculate_provider_metric_score', 6.44594031077735, 276320], ['ZipHealth.get_csm_registries', 14.00570643953125, 12800], ['ZipHealth.ccd_bulk_review', 36.272532497214684, 7001], ['ZipHealth.ccda_member_matching_event', 128.98130220909098, 7810], ['ZipHealth.get_organisation_list_for_drop_down', 5.81742264581231, 1982], ['ZipHealth.get_API_details', 12.775653618481495, 3701], ['ZipHealth.user_email_notification', 3.443967, 1], ['ZipHealth.patient_details', 4.99387256052363, 5882], ['ZipHealth.get_banner', 11.62860328400196, 4088], ['ZipHealth.calculate_care_gap_list_for_org_hcc', 19.044961148047538, 2945], ['ZipHealth.manage_chart_list_export', 35.020795609970676, 2046], ['ZipHealth._care_opp_details', 4.274826643389565, 2549], ['ZipHealth.set_contact_type_email', 3.3800734584837544, 554], ['ZipHealth.get_support_activity', 3.0242969139273437, 44567]]

launch_gui()
#generate_histogram_graph(test_data)
count =0
for i in querytime:
    if i>30:
        count+=1
#print(count)




