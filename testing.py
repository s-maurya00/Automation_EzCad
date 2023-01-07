import time
from pywinauto.application import Application
from pywinauto.keyboard import send_keys

# Open the EzCad application
app = Application(backend='uia').start('C:\\Users\\maury\\Desktop\\For OptiLOM\\EZCAD2-Software\\EZCAD2 For AiO (20220915 Release)\\EZCAD2 For AiO.exe')

# # app.EzCad.print_control_identifiers()     # Used to Identify the ID of the element you want to perform the action with
Error1 = app.EzCad.child_window(title="OK", auto_id="2", control_type="Button").wrapper_object()    # Create the reference of the object element
Error1.click_input()    # Function to click the object element

Error1 = app.EzCad.child_window(title="OK", auto_id="2", control_type="Button").wrapper_object()
Error1.click_input()

time.sleep(3)
FileMenu = app.EzCad21411NoTitle.child_window(title="File", control_type="MenuItem").wrapper_object()
FileMenu.click_input()

# # app.EzCad21411NoTitle.print_control_identifiers()
ImportVectorFile = app.EzCad21411NoTitle.child_window(title="Import Vector File...\tCtrl+B", auto_id="32807", control_type="MenuItem").wrapper_object()
ImportVectorFile.click()


# app.Open.print_control_identifiers()



# """By following Rafique Javed's youtube video"""
# # For error section
# app.EzCad.ok.click()
# app.EzCad.ok.click()

# # For importing file
# time.sleep(2)
# app.EzCad21411NoTitle.MenuItem(u'&File->&Import Vector File').click()