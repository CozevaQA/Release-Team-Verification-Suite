try:
    import variablestorage
except IndexError as e:
    import FirstTimeSetup

import guiwindow
import setups
import ExcelProcessor as db
import logging
if __name__ == '__main__':
    x=0
current_client_id = '9999'
try:
    with open("assets\\overwatch_cache.txt", 'r+') as file:
        content = file.read().strip()
        if content.isdigit():
            number = int(content)
            current_client_id = str(number).strip()
            file.seek(0)
            file.truncate()
        else:
            print("File does not contain a valid number.")
except IOError as e:
    print(f"Error: {e}")



privacy_status = 'Onshore'
roleset = db.getDefaultUserNames(db.fetchCustomerName(current_client_id))
if "Customer Support" in roleset:
    privacy_status = 'Offshore'
feature_checklist = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

# comment the next 3 lines out when you need to run default. DO NOT COMMIT THIS YOU WILL BREAK EVERYTHING
#feature_checklist = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
#roleset = {'Cozeva Support': '99999'}
#privacy_status = 'Onshore'
guiwindow.verification_specs = [db.fetchCustomerName(current_client_id), current_client_id, privacy_status, roleset, feature_checklist]


print(guiwindow.verification_specs)
print("Enviromnent: "+guiwindow.env)
print("Headless Mode: "+str(guiwindow.headlessmode))
environment = guiwindow.env

launchstyle= "Def"
if 'NC_' in guiwindow.verification_specs[0]:
    launchstyle = "NC"
print(launchstyle)
if guiwindow.verification_specs[0] == 'Name' or guiwindow.verification_specs[0] == 'Customer' and guiwindow.verification_specs[4][13] == 0:
    exit(9)
if guiwindow.verification_specs[4][13] == 1:
    import runner
    exit(45)
driver_created = 0
setups.driver_setup()
driver_created = 1
if environment == "PROD":
    setups.login_to_cozeva(guiwindow.verification_specs[1])
elif environment == "CERT":
    setups.login_to_cozeva_cert(guiwindow.verification_specs[1])
if guiwindow.verification_specs[2] == "Onshore":
    if launchstyle == "Def":
        setups.cozeva_support(environment)
    elif launchstyle == "NC":
        setups.new_launch(environment)
elif guiwindow.verification_specs[2] == "Offshore":
    roleset = guiwindow.verification_specs[3]
    for roles in roleset:
        if roles == "Cozeva Support":
            print("skipping Cozeva Support because its offshore customer")
        elif roles == "Limited Cozeva Support":
            print("Run Limited Cozeva Support Verification for username " + roleset[roles])
            setups.limited_cozeva_support(roleset[roles])
        elif roles == "Customer Support":
            print("Run Customer Support Verification for username " + roleset[roles])
            setups.customer_support(roleset[roles])
        elif roles == "Regional Support":
            print("Run Regional Support Verification for username " + roleset[roles])
            setups.regional_suport(roleset[roles])
        elif roles == "Office Admin Practice Delegate":
            print("Run Office Admin Practice Delegate Verification for username " + roleset[roles])
            setups.office_admin_Prac(roleset[roles])
        elif roles == "Office Admin Provider Delegate":
            print("Run Office Admin Provider Delegate Verification for username " + roleset[roles])
            setups.office_admin_prov(roleset[roles])
        elif roles == "Provider":
            print("Run Provider Verification for username " + roleset[roles])
            setups.prov(roleset[roles])

if driver_created == 1:
    setups.driver.quit()