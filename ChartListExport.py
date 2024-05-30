# C:\Users\psur\AppData\Local\Programs\Python\Python311\python.exe "G:\My Drive\IMPORTANT\Python Selenium Scripts\Pritam\ChartListExport.py"
import pandas as pd
import re

# customer_id = 9999
# file_path_supplemental_data = 'C:/Users/username/Downloads/Supplemental Data 2024-04-16.csv'
# file_path_hcc_chart_list = 'C:/Users/username/Downloads/HCC Chart List 2024-04-16.csv'
# file_path_awv_chart_list = 'C:/Users/username/Downloads/AWV Chart List 2024-04-12 (2).csv'
# file_path_report = 'C:/Users/username/Downloads/'


def supplemental_data(customer_id, file_path_supplemental_data, report_body, observations_body):
    if file_path_supplemental_data != 0:
        print("\n**** SUPPLEMENTAL DATA LISTS ****")
        df = pd.read_csv(file_path_supplemental_data)
        if customer_id == 200:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
            report_body += "<h2>Customer: HPMG </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: HPMG </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 3000:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "PPG", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
            report_body += "<h2>Customer: Health Net </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: Health Net </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 4600:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
            report_body += "<h2>Customer: MedPOINT Management </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: MedPOINT Management </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 3300:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Affiliated Provider's Practice", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date",
                                "Review 2 by", "Review 2 comment", "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status",
                                "Bridge source", "Enrollment Start", "Enrollment End"]
            report_body += "<h2>Customer: OPTUM </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: OPTUM </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 1300:
            expected_columns = ["Task #","Cozeva ID","Patient Name","Gender","DOB","Member ID","Member UID","Health Plan","Product Code","Service Date","Rendering / Reviewing Provider","Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID","Rendering / Reviewing Provider NPI","Measure","Code & Value","Excluded","Status","Status Reason","Affiliated Provider","Affiliated Provider ID","Affiliated Provider UID",
                                "Submitter status","Submitter date","Submitted by","Submitter comment","Review 1 status","Review 1 date","Review 1 by","Review 1 comment","Review 2 status","Review 2 date","Review 2 by","Review 2 comment",
                                "Review 3 status","Review 3 date","Review 3 by","Review 3 comment","Created On"," Created by","Labels","Measurement Year","Submission Type","Group Name","Link Status","Bridge source","Enrollment Start","Enrollment End"]
            report_body += "<h2>Customer: Prospect </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: Prospect </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 6800:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start",
                                "Enrollment End"]
            report_body += "<h2>Customer: Brand New Day </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: Brand New Day </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 6700:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start",
                                "Enrollment End"]
            report_body += "<h2>Customer: Central Health Plan </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: Central Health Plan </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 1850:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
            report_body += "<h2>Customer: Providence CA </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: Providence CA </h2><h3>Supplemental Data List:</h3>"
        elif customer_id == 1000:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Created On", " Created by", "Labels", "Measurement Year",
                                "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
            report_body += "<h2>Customer: SCPMCS </h2><h3>Supplemental Data List:</h3>"
            observations_body += "<h2>Customer: SCPMCS </h2><h3>Supplemental Data List:</h3>"
        else:
            expected_columns = []
            report_body += "<h2>Customer: Need to be Configured </h2><h3>Supplemental Data List:</h3>"
        # Missing Column(s) Check
        report_body += "<ul>"
        observations_body += "<ul>"
        missing_columns = set(expected_columns) - set(df.columns)
        if not missing_columns:
            print("All columns are present.")
            report_body += "<li>All columns are present.</li><p></p>"
        else:
            print("Missing columns: {}".format(', '.join(missing_columns)))
            missing_columns_text = "Missing columns: {}".format(', '.join(missing_columns))
            report_body += "<li><b style='color:red'> " + missing_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + missing_columns_text + "</li><p></p>"
        # Newly Added Column(s) Check
        newly_added_columns = set(df.columns) - set(expected_columns)
        if not newly_added_columns:
            print("No new columns added.")
            report_body += "<li>No new columns added.</li><p></p>"
        else:
            print("Newly added columns: {}".format(', '.join(newly_added_columns)))
            newly_added_columns_text = "Newly added columns: {}".format(', '.join(newly_added_columns))
            report_body += "<li><b style='color:red'> " + newly_added_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + newly_added_columns_text + "</li><p></p>"
        # Blank Columns Check
        blank_columns = df.columns[df.isna().all()]
        if len(blank_columns) > 0:
            print(f"The following columns contain all blank values: {', '.join(blank_columns)}")
            blank_columns_text = f"Column(s) contain all blank values: {', '.join(blank_columns)}"
            report_body += "<li><b style='color:red'> " + blank_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + blank_columns_text + "</li><p></p>"
        else:
            print("No columns contain all blank values.")
            report_body += "<li>No columns contain all blank values.</li><p></p>"
        # Check if the column contains only numbers
        columns_to_check = ['Measurement Year']
        numeric_results = {}
        for column in columns_to_check:
            is_numeric = pd.to_numeric(df[column], errors='coerce').notna().all()
            numeric_results[column] = is_numeric
        for column, result in numeric_results.items():
            if result:
                print(f"All values in column '{column}' are numeric.")
                report_body += "<li>" + f"All values in column '{column}' are numeric." + "</li><p></p>"
            else:
                print(f"Not all values in column '{column}' are numeric.")
                report_body += "<li><b style='color:red'> " + f"Not all values in column '{column}' are numeric." + "</b></li><p></p>"
                observations_body += "<li>" + f"Not all values in column '{column}' are numeric." + "</li><p></p>"
        # Check if the column values are in Date format
        expected_formats = [
            re.compile(r'^\d{2}-\d{2}-\d{4}$'),  # Format: XX-XX-XXXX
            re.compile(r'^\d{2}/\d{2}/\d{4}$'),  # Format: XX/XX/XXXX
            ]
        columns_to_check = ['DOB', 'Service Date', 'Submitter date']
        for column_to_check in columns_to_check:
            # Check if all values in the column match any of the expected formats
            valid_strings = df[column_to_check].apply(lambda x: isinstance(x, str))
            are_all_matching = valid_strings.all() and df[column_to_check][valid_strings].apply(lambda x: any(pattern.match(x) for pattern in expected_formats)).all()
            if are_all_matching:
                print(f"All values in column '{column_to_check}' match the expected formats.")
                report_body += "<li>" + f"All values in column '{column_to_check}' match the expected formats." + "</li><p></p>"
            else:
                print(f"Not all values in column '{column_to_check}' match the expected formats.")
                report_body += "<li><b style='color:red'> " + f"Not all values in column '{column_to_check}' match the expected formats." + "</b></li><p></p>"
                observations_body += "<li>" + f"Not all values in column '{column_to_check}' match the expected formats." + "</li><p></p>"
        report_body += "</ul>"
        observations_body += "</ul>"
        return report_body, observations_body
    else:
        report_body += ""
        observations_body += ""
        return report_body, observations_body


def hcc_chart_list(customer_id, file_path_hcc_chart_list, report_body, observations_body):
    if file_path_hcc_chart_list != 0:
        print("\n**** HCC CHART LISTS ****")
        df = pd.read_csv(file_path_hcc_chart_list)
        report_body += "<h3>HCC Chart List:</h3>"
        observations_body += "<h3>HCC Chart List:</h3>"
        if customer_id == 200:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 3000:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "PPG", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 4600:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Pre-review status", "Pre-review date", "Pre-review by", "Pre-review comment", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by",
                                "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source",
                                "Enrollment Start", "Enrollment End"]
        elif customer_id == 1300:
            expected_columns = ["Task #","Cozeva ID","Patient Name","Gender","DOB","Member ID","Member UID","Health Plan","Product Code","Service Date","Rendering / Reviewing Provider","Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID","Rendering / Reviewing Provider NPI","Condition","Code & Value","Excluded","Status","Status Reason","Affiliated Provider","Affiliated Provider ID","Affiliated Provider UID",
                                "Pre-review status","Pre-review date","Pre-review by","Pre-review comment","Submitter status","Submitter date","Submitted by","Submitter comment","Review 1 status","Review 1 date","Review 1 by","Review 1 comment",
                                "Review 2 status","Review 2 date","Review 2 by","Review 2 comment","Review 3 status","Review 3 date","Review 3 by","Review 3 comment","Created On"," Created by","Labels","Measurement Year","Submission Type","Group Name",
                                "Link Status","Bridge source","Enrollment Start","Enrollment End"]
        elif customer_id == 6800:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start",
                                "Enrollment End"]
        elif customer_id == 6700:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment",
                                "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start",
                                "Enrollment End"]
        elif customer_id == 1850:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Pre-review status", "Pre-review date", "Pre-review by", "Pre-review comment", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by",
                                "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source",
                                "Enrollment Start", "Enrollment End"]
        elif customer_id == 1000:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID", "Affiliated Provider UID",
                                "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Created On", " Created by", "Labels", "Measurement Year",
                                "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
        else:
            expected_columns = []
        # Missing Column(s) Check
        report_body += "<ul>"
        observations_body += "<ul>"
        missing_columns = set(expected_columns) - set(df.columns)
        if not missing_columns:
            print("All columns are present.")
            report_body += "<li>All columns are present.</li><p></p>"
        else:
            print("Missing columns: {}".format(', '.join(missing_columns)))
            missing_columns_text = "Missing columns: {}".format(', '.join(missing_columns))
            report_body += "<li><b style='color:red'> " + missing_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + missing_columns_text + "</li><p></p>"
        # Newly Added Column(s) Check
        newly_added_columns = set(df.columns) - set(expected_columns)
        if not newly_added_columns:
            print("No new columns added.")
            report_body += "<li>No new columns added.</li><p></p>"
        else:
            print("Newly added columns: {}".format(', '.join(newly_added_columns)))
            newly_added_columns_text = "Newly added columns: {}".format(', '.join(newly_added_columns))
            report_body += "<li><b style='color:red'> " + newly_added_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + newly_added_columns_text + "</li><p></p>"
        # Blank Columns Check
        blank_columns = df.columns[df.isna().all()]
        if len(blank_columns) > 0:
            print(f"The following columns contain all blank values: {', '.join(blank_columns)}")
            blank_columns_text = f"Column(s) contain all blank values: {', '.join(blank_columns)}"
            report_body += "<li><b style='color:red'> " + blank_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + blank_columns_text + "</li><p></p>"
        else:
            print("No columns contain all blank values.")
            report_body += "<li>No columns contain all blank values.</li><p></p>"
        # Check if the column contains only numbers
        columns_to_check = ['Measurement Year']
        numeric_results = {}
        for column in columns_to_check:
            is_numeric = pd.to_numeric(df[column], errors='coerce').notna().all()
            numeric_results[column] = is_numeric
        for column, result in numeric_results.items():
            if result:
                print(f"All values in column '{column}' are numeric.")
                report_body += "<li>" + f"All values in column '{column}' are numeric." + "</li><p></p>"
            else:
                print(f"Not all values in column '{column}' are numeric.")
                report_body += "<li><b style='color:red'> " + f"Not all values in column '{column}' are numeric." + "</b></li><p></p>"
                observations_body += "<li>" + f"Not all values in column '{column}' are numeric." + "</li><p></p>"
        # Check if the column values are in Date format
        expected_formats = [
            re.compile(r'^\d{2}-\d{2}-\d{4}$'),  # Format: XX-XX-XXXX
            re.compile(r'^\d{2}/\d{2}/\d{4}$'),  # Format: XX/XX/XXXX
            ]
        columns_to_check = ['DOB', 'Service Date', 'Submitter date']
        for column_to_check in columns_to_check:
            # Check if all values in the column match any of the expected formats
            valid_strings = df[column_to_check].apply(lambda x: isinstance(x, str))
            are_all_matching = valid_strings.all() and df[column_to_check][valid_strings].apply(lambda x: any(pattern.match(x) for pattern in expected_formats)).all()
            if are_all_matching:
                print(f"All values in column '{column_to_check}' match the expected formats.")
                report_body += "<li>" + f"All values in column '{column_to_check}' match the expected formats." + "</li><p></p>"
            else:
                print(f"Not all values in column '{column_to_check}' match the expected formats.")
                report_body += "<li><b style='color:red'> " + f"Not all values in column '{column_to_check}' match the expected formats." + "</b></li><p></p>"
                observations_body += "<li>" + f"Not all values in column '{column_to_check}' match the expected formats." + "</li><p></p>"
        report_body += "</ul>"
        observations_body += "</ul>"
        return report_body, observations_body
    else:
        report_body += ""
        observations_body += ""
        return report_body, observations_body


def awv_chart_list(customer_id, file_path_awv_chart_list, report_body, observations_body):
    if file_path_awv_chart_list != 0:
        print("\n**** AWV CHART LISTS ****")
        df = pd.read_csv(file_path_awv_chart_list)
        report_body += "<h3>AWV Chart List:</h3>"
        observations_body += "<h3>AWV Chart List:</h3>"
        if customer_id == 200:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date",
                                "Review 2 by", "Review 2 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 3000:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "PPG", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date",
                                "Review 2 by", "Review 2 comment", "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Link Status", "Bridge source",
                                "Enrollment Start", "Enrollment End"]
        elif customer_id == 4600:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Pre-review status", "Pre-review date", "Pre-review by", "Pre-review comment", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date",
                                "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status",
                                "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 1300:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Pre-review status", "Pre-review date", "Pre-review by", "Pre-review comment", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date",
                                "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date", "Review 2 by", "Review 2 comment", "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels",
                                "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 6800:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date",
                                "Review 2 by", "Review 2 comment", "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status",
                                "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 6700:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Review 2 status", "Review 2 date",
                                "Review 2 by", "Review 2 comment", "Review 3 status", "Review 3 date", "Review 3 by", "Review 3 comment", "Created On", " Created by", "Labels", "Measurement Year", "Submission Type", "Group Name", "Link Status",
                                "Bridge source", "Enrollment Start", "Enrollment End"]
        elif customer_id == 1000:
            expected_columns = ["Task #", "Cozeva ID", "Patient Name", "Gender", "DOB", "Member ID", "Member UID", "Health Plan", "Product Code", "Service Date", "Rendering / Reviewing Provider", "Rendering / Reviewing Provider ID",
                                "Rendering / Reviewing Provider UID", "Rendering / Reviewing Provider NPI", "Measure / Condition", "Code & Value", "Excluded", "Status", "Status Reason", "Affiliated Provider", "Affiliated Provider ID",
                                "Affiliated Provider UID", "Submitter status", "Submitter date", "Submitted by", "Submitter comment", "Review 1 status", "Review 1 date", "Review 1 by", "Review 1 comment", "Created On", " Created by", "Labels",
                                "Measurement Year", "Submission Type", "Group Name", "Link Status", "Bridge source", "Enrollment Start", "Enrollment End"]
        else:
            expected_columns = []
        # Missing Column(s) Check
        report_body += "<ul>"
        observations_body += "<ul>"
        missing_columns = set(expected_columns) - set(df.columns)
        if not missing_columns:
            print("All columns are present.")
            report_body += "<li>All columns are present.</li><p></p>"
        else:
            print("Missing columns: {}".format(', '.join(missing_columns)))
            missing_columns_text = "Missing columns: {}".format(', '.join(missing_columns))
            report_body += "<li><b style='color:red'> " + missing_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + missing_columns_text + "</li><p></p>"
        # Newly Added Column(s) Check
        newly_added_columns = set(df.columns) - set(expected_columns)
        if not newly_added_columns:
            print("No new columns added.")
            report_body += "<li>No new columns added.</li><p></p>"
        else:
            print("Newly added columns: {}".format(', '.join(newly_added_columns)))
            newly_added_columns_text = "Newly added columns: {}".format(', '.join(newly_added_columns))
            report_body += "<li><b style='color:red'> " + newly_added_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + newly_added_columns_text + "</li><p></p>"
        # Blank Columns Check
        blank_columns = df.columns[df.isna().all()]
        if len(blank_columns) > 0:
            print(f"The following columns contain all blank values: {', '.join(blank_columns)}")
            blank_columns_text = f"Column(s) contain all blank values: {', '.join(blank_columns)}"
            report_body += "<li><b style='color:red'> " + blank_columns_text + "</b></li><p></p>"
            observations_body += "<li>" + blank_columns_text + "</li><p></p>"
        else:
            print("No columns contain all blank values.")
            report_body += "<li>No columns contain all blank values.</li><p></p>"
        # Check if the column contains only numbers
        columns_to_check = ['Measurement Year']
        numeric_results = {}
        for column in columns_to_check:
            is_numeric = pd.to_numeric(df[column], errors='coerce').notna().all()
            numeric_results[column] = is_numeric
        for column, result in numeric_results.items():
            if result:
                print(f"All values in column '{column}' are numeric.")
                report_body += "<li>" + f"All values in column '{column}' are numeric." + "</li><p></p>"
            else:
                print(f"Not all values in column '{column}' are numeric.")
                report_body += "<li><b style='color:red'> " + f"Not all values in column '{column}' are numeric." + "</b></li><p></p>"
                observations_body += "<li>" + f"Not all values in column '{column}' are numeric." + "</li><p></p>"
        # Check if the column values are in Date format
        expected_formats = [
            re.compile(r'^\d{2}-\d{2}-\d{4}$'),  # Format: XX-XX-XXXX
            re.compile(r'^\d{2}/\d{2}/\d{4}$'),  # Format: XX/XX/XXXX
            ]
        columns_to_check = ['DOB', 'Service Date', 'Submitter date']
        for column_to_check in columns_to_check:
            # Check if all values in the column match any of the expected formats
            valid_strings = df[column_to_check].apply(lambda x: isinstance(x, str))
            are_all_matching = valid_strings.all() and df[column_to_check][valid_strings].apply(lambda x: any(pattern.match(x) for pattern in expected_formats)).all()
            if are_all_matching:
                print(f"All values in column '{column_to_check}' match the expected formats.")
                report_body += "<li>" + f"All values in column '{column_to_check}' match the expected formats." + "</li><p></p>"
            else:
                print(f"Not all values in column '{column_to_check}' match the expected formats.")
                report_body += "<li><b style='color:red'> " + f"Not all values in column '{column_to_check}' match the expected formats." + "</b></li><p></p>"
                observations_body += "<li>" + f"Not all values in column '{column_to_check}' match the expected formats." + "</li><p></p>"
        report_body += "</ul>"
        observations_body += "</ul>"
        return report_body, observations_body
    else:
        report_body += ""
        observations_body += ""
        return report_body, observations_body


def main_chart_list_export(customer_id, file_path_supplemental_data, file_path_hcc_chart_list, file_path_awv_chart_list, file_path_report):
    report_body = "<h2> Chart List Export Verification Report: </h2>"
    observations_body = "<h2> Issues/Observations: </h2>"
    report_body, observations_body = supplemental_data(customer_id, file_path_supplemental_data, report_body, observations_body)
    report_body, observations_body = hcc_chart_list(customer_id, file_path_hcc_chart_list, report_body, observations_body)
    report_body, observations_body = awv_chart_list(customer_id, file_path_awv_chart_list, report_body, observations_body)
    report_html_content = """<!DOCTYPE html><html><head><title>Chart List Report</title></head><body>""" + report_body + """</body></html>"""
    with open(file_path_report + str(customer_id) + '_Reports.html', 'w') as f:
        f.write(report_html_content)
    observation_html_content = """<!DOCTYPE html><html><head><title>Observations</title></head><body>""" + observations_body + """</body></html>"""
    with open(file_path_report + str(customer_id) + '_Observations.html', 'w') as f:
        f.write(observation_html_content)


# def main():
#     main_chart_list_export(customer_id, file_path_supplemental_data, file_path_hcc_chart_list, file_path_awv_chart_list, file_path_report)


# if __name__ == "__main__":
#     main()