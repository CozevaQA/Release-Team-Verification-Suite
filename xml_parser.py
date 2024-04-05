import configparser
import xml.etree.ElementTree as ET
from openpyxl import Workbook, load_workbook
from tkinter import *
from os import listdir
from os.path import isfile, join

config = configparser.RawConfigParser()
config.read("templatemap.properties")
path = "C:\\VerificationReports\\"
titles = []
number_of_entries = []
template_id = []
table_names = []
queries = []

status_code_map = {"active": "11",
                   "completed": "13",
                   "canceled": "15",
                   "aborted": "10",
                   "new": "32",
                   "request": "16",
                   "requested": "16",
                   "inactive": "31",
                   "resolved": "7",
                   "suspended": "12"}


def parse_xml(xml_path, ccd_id, filename):
    tree = ET.parse(xml_path)
    xml_root = tree.getroot()
    sections = xml_root.find("{urn:hl7-org:v3}component").find("{urn:hl7-org:v3}structuredBody").findall(
        "{urn:hl7-org:v3}component")
    for component in sections:
        titles.append(component.find("{urn:hl7-org:v3}section").find("{urn:hl7-org:v3}title").text)
    for component in sections:
        number_of_entries.append(len(component.find("{urn:hl7-org:v3}section").findall("{urn:hl7-org:v3}entry")))
    for component in sections:
        temp_list = []
        id_list = component.find("{urn:hl7-org:v3}section").findall("{urn:hl7-org:v3}templateId")
        for id_element in id_list:
            temp_list.append(id_element.get("root"))
        template_id.append(sorted([*set(temp_list)], key=len, reverse=True))

    wb = Workbook()
    ws = wb.active

    ws.append(["Component title", "Number of Entries", "Longest Template ID", "Additional Template IDs", "Table Name",
               "Query Set"])

    for title, entry, template in zip(titles, number_of_entries, template_id):
        print(title + "\t" + str(entry) + "\t" + str(template))
        try:
            table_name = config.get("templateids", str(template[:1][0]))
            query_name = "SELECT * FROM " + str(table_name) + " WHERE ccd_id = " + ccd_id + ";"
            ws.append([title, str(entry), str(template[:1][0]), str(template[1:]), table_name, query_name])
            table_names.append(table_name)
            queries.append(query_name)
        except IndexError as e:
            ws.append([title, str(entry), "Template ID missing", "-", ""])
            table_names.append("-")
            queries.append("-")
        except configparser.NoOptionError as e:
            # look in the other templateids
            found_flag = 0
            if len(template[1:]) > 0:
                for id in template[1:]:
                    try:
                        table_name = config.get("templateids", str(id))
                        found_flag = 1
                    except configparser.NoOptionError as e:
                        continue
            if found_flag == 1:
                query_name = "SELECT * FROM " + str(table_name) + " WHERE ccd_id = " + ccd_id + ";"
                ws.append([title, str(entry), str(template[:1][0]), str(template[1:]), table_name, query_name])
                table_names.append(table_name)
                queries.append(query_name)
            else:
                ws.append(
                    [title, str(entry), str(template[:1][0]), str(template[1:]), "No matching templateID in map", ""])
                table_names.append("-")
                queries.append("-")
    export_path = path + filename.replace(".xml", "_") + ccd_id + ".xlsx"
    wb.save(export_path)


def parse_specific_section(xml_path, ccd_id, filename):  # medication_dispense_only
    tree = ET.parse(xml_path)
    xml_root = tree.getroot()
    medication_template = config.get("templatemap", "medications_section")
    medication_component = None
    sections = xml_root.find("{urn:hl7-org:v3}component").find("{urn:hl7-org:v3}structuredBody").findall(
        "{urn:hl7-org:v3}component")
    # extract the relevant medication section into a variable
    section_found = 0
    for component in sections:
        # create a template list of all the template IDS in the section
        template_tags = component.find("{urn:hl7-org:v3}section").findall("{urn:hl7-org:v3}templateId")
        for template_root in template_tags:
            root_value = template_root.get('root')
            if root_value == medication_template:
                section_found = 1
                break
        if section_found == 1:
            medication_component = component
            break
        else:
            continue

    if section_found == 0:
        print("No medication section")
        return
    # when the component is found, load the entries into a list.
    medication_entries = medication_component.find("{urn:hl7-org:v3}section").findall("{urn:hl7-org:v3}entry")
    # loop through the entries and store all medication dispense specific tags in one list.
    medication_dispense = []

    for entry in medication_entries:
        entry_relationship_tags = entry.find("{urn:hl7-org:v3}substanceAdministration").findall(
            "{urn:hl7-org:v3}entryRelationship")
        for potential_dispense in entry_relationship_tags:
            if len(potential_dispense.findall("{urn:hl7-org:v3}supply")) > 0:
                if potential_dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}templateId").get(
                        'root') == config.get("templatemap", "medication_dispense"):
                    medication_dispense.append(potential_dispense)

    # Now, extract relevant data from the dispenses and display one by one.
    for index, dispense in enumerate(medication_dispense):
        try:
            med_info_product_code = dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}product").find(
                "{urn:hl7-org:v3}manufacturedProduct").find("{urn:hl7-org:v3}manufacturedMaterial").find(
                "{urn:hl7-org:v3}code").get('code')
            med_info_product_code_desc = dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}product").find(
                "{urn:hl7-org:v3}manufacturedProduct").find("{urn:hl7-org:v3}manufacturedMaterial").find(
                "{urn:hl7-org:v3}code").get('displayName')
            status_id = status_code_map[
                dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}statusCode").get('code')]
            repeat_no = dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}repeatNumber").get('value')
            dose_value = dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}quantity").get('value')
            time_offset = dispense.find("{urn:hl7-org:v3}supply").find("{urn:hl7-org:v3}effectiveTime").get('value')
            print(med_info_product_code, med_info_product_code_desc, status_id, repeat_no, dose_value, time_offset)

        except AttributeError as e:
            print("Dispense # " + str(index + 1) + " is not usable")


def master_parser_gui():
    # design - 2 entry fields (XML name and CCD ID) which will get passed and called for processing. call the
    # parse_xml() function from inside here on button press
    root = Tk()

    xml_parent_path = "C:\\xmlfiles"
    xml_filename = ""

    def parse_button():
        parse_xml(join(xml_parent_path, selected_file.get()), ccd_id_entry.get(), selected_file.get())
        root.destroy()
        master_parser_result_gui()

    def parse_section():
        parse_specific_section(join(xml_parent_path, selected_file.get()), ccd_id_entry.get(), selected_file.get())
        root.destroy()

    # xml_filename_entry = Entry(root)
    ccd_id_entry = Entry(root)

    selected_file = StringVar()
    selected_file.set("Select...")
    xml_list = [f for f in listdir(xml_parent_path) if isfile(join(xml_parent_path, f))]  # vs.customer_list
    xml_drop = OptionMenu(root, selected_file, *xml_list)

    Label(root, text="Enter XML filename").grid(row=0, column=2, sticky='w')
    xml_drop.grid(row=1, column=2, sticky='w')
    Label(root, text="Enter CCD ID in UI").grid(row=0, column=4, sticky='w')
    ccd_id_entry.grid(row=1, column=4, sticky='w')
    Button(root, text="Begin Processing", command=parse_button).grid(row=2, column=3, sticky='w')
    Button(root, text="Medication dispense", command=parse_section).grid(row=3, column=3, sticky='w')

    # xml_filename_entry.insert(0, "epic_hill_4.xml")
    ccd_id_entry.insert(0, "1234")

    root.title("XML Master Parser")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()


def master_parser_result_gui():
    root = Tk()

    # form column names
    Label(root, text="Section Name").grid(row=0, column=0, sticky='w')
    Label(root, text="# of Entries").grid(row=0, column=1, sticky='w')
    Label(root, text="Query Set").grid(row=0, column=2, sticky='w')
    querybox = Text(root, height=len(number_of_entries))
    query_string = """"""
    row_counter = 1
    query_counter = 0
    for title, entry, query in zip(titles, number_of_entries, queries):
        Label(root, text=title).grid(row=row_counter, column=0, sticky='w')
        Label(root, text=entry).grid(row=row_counter, column=1, sticky='w')
        if query != "-":
            query_string += query
            query_counter += 1
            if entry == 0:
                query_string += " #No entries found for this section"
            query_string += "\n"
        else:
            query_string += ""


        row_counter += 1
    querybox.insert(INSERT, query_string)
    querybox.config(height=row_counter - 1 - (row_counter - 1 - query_counter))
    querybox.grid(row=1, rowspan=row_counter - 1, column=2, columnspan=4)

    root.title("XML Master Parser Results")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()


master_parser_gui()
