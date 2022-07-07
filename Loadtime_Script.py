import setups
import logging
import ExcelProcessor as db
import context_functions as cf
import support_functions as sf
if __name__ == '__main__':
    print("Hello World")
    driver = setups.driver_setup()
    setups.login_to_cozeva()
    customer_list=db.getCustomerList()
    path = setups.create_folders("Cozeva Support")
    workbook = setups.create_reporting_workbook(path)
    for cust in customer_list:
        id=db.fetchCustomerID(cust)
        setups.switch_customer_context(id)
        workbook.save(path + '\\Report.xlsx')
        cf.click_on_each_metric(cust, driver, workbook, path)
        workbook.save(path + "\\Report.xlsx")






