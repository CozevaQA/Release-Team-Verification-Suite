

try:
    import variablestorage
except IndexError as e:
    import FirstTimeSetup

# Aritra Da Push test


import guiwindow
import setups
import logging
if __name__ == '__main__':
    print("Hello World")
    guiwindow.launchgui()
    print(guiwindow.verification_specs)
    if guiwindow.verification_specs[0] == 'Name' or guiwindow.verification_specs[0] == 'Customer' and guiwindow.verification_specs[4][13] == 0:
        exit(9)
    if guiwindow.verification_specs[4][13] == 1:
        import runner
        #ugh how this work
        #testing testing


        exit(45)
    driver_created = 0
    setups.driver_setup()
    driver_created = 1
    setups.login_to_cozeva()
    if guiwindow.verification_specs[2] == "Onshore":
        setups.cozeva_support()
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









