import time
from pywinauto.application import Application
from pywinauto.keyboard import send_keys


def selectFirstObjectInList(EzCadAppRef):
    # Selects the first item to be printed from the available objects in the object list

    # Select the "Edit Node" Tool
    toolbarwindow = EzCadAppRef[u'Toolbar3']
    edit_node = toolbarwindow.button(1)
    edit_node.click()

    # Reselect the "Pick" Tool
    pick = toolbarwindow.button(0)
    pick.click()

    # Click the title of "Object List"
    objectHeader = EzCadAppRef[u'Header']
    objectHeader.click()
    # "Home" key selects the 1st element in the object list
    send_keys('{HOME}')


def setXtoZero(EzCadAppRef):
    # Sets the X axis position of current object to "0" and applies it

    EzCadAppRef[u'PositionEdit'].type_keys('{END}+{HOME}0')
    EzCadAppRef[u'&Apply'].click_input()


def hatchObject(app, EzCadAppRef):
    # Hatches the current object and press 'OK'

    EzCadAppRef.menu_item(u'&Edit->Hatch\\tCtrl+H').click()
    app.Hatch[u'&OK'].click_input()


def selectMarkingProperty(EzCadAppRef, iter):
    # Selects either 'black' or 'blue' marking property based on the 'iter' value

    if(iter == 0):
        EzCadAppRef[u'Button12'].click_input()
    elif(iter == 1):
        EzCadAppRef[u'Button13'].click_input()


# INCOMPLETE
def startMarking(EzCadAppRef):
    # Start the Marking process

    EzCadAppRef[u'Mark(F2)'].click_input()


def clickEnableInHatching(EzCadAppRef):
    # Enable OR Disable the hatching checkbox in the "Hatching" Table and click "Apply"

    EzCadAppRef[u'Enable'].click_input()
    EzCadAppRef[u'&Apply'].click_input()


def print3dItem(app, EzCadAppRef):
    # Performs all relevant steps for all the available objects

    for i in range(10):
        selectFirstObjectInList(EzCadAppRef)

        time.sleep(0.5)
        setXtoZero(EzCadAppRef)

        time.sleep(0.5)
        hatchObject(app, EzCadAppRef)

        time.sleep(0.5)
        selectMarkingProperty(EzCadAppRef, 0)

        time.sleep(0.5)
        startMarking(EzCadAppRef)
        time.sleep(2)   # Sleep till the printing is completed

        time.sleep(0.5)
        clickEnableInHatching(EzCadAppRef)

        time.sleep(0.5)
        selectMarkingProperty(EzCadAppRef, 1)

        time.sleep(0.5)
        startMarking(EzCadAppRef)
        time.sleep(2)

        time.sleep(0.5)
        clickEnableInHatching(EzCadAppRef)

        time.sleep(0.5)
        send_keys('{DELETE}')





app = Application().start(cmd_line=u'"C:\\Users\\maury\\Desktop\\For OptiLOM\\EZCAD2-Software\\EZCAD2 For AiO (20220915 Release)\\EZCAD2 For AiO.exe" ')
EzCadAppRef = app[u'EzCad2.14.11 - No title']


# Handling Errors
errorWindow = app.EzCad.Ok.click()
errorWindow = app.EzCad.Ok.click()


# Pause code's execuition untill the application gets loaded
EzCadAppRef.wait('ready')


time.sleep(1)
# Open the window to import file
EzCadAppRef.menu_item(u'&File->Import Vector File...\\tCtrl+B').click()


# app -> has "Open" window -> has "Edit" tab -> and then we click it
### window = app.Open


# Click on "File Name" input Box
### edit = window.Edit
### edit.click()


# Type the URL of file in input-box
app.Open.Edit.type_keys("C:\\Users\\maury\\Desktop\\For OptiLOM\\testingInputForEZCAD\\Untitled.svg", with_spaces = True)


time.sleep(1)
# Click 'OK' button in "Open" window
app.Open[u'&Open'].click_input()


# Click and then Set the checkbox position to 2nd row, 1st Column
EzCadAppRef.Button3.click()
app.Dialog[u'CheckBox8'].click()
app.Dialog.Ok.click()


# Set 'X axis position', 'Y axis position' to Zero and then click 'Apply' button
EzCadAppRef.Edit2.type_keys('{END}+{HOME}0')
EzCadAppRef.Edit3.type_keys('{END}+{HOME}0')
EzCadAppRef[u'&Apply'].click()


# Ungroup elements
EzCadAppRef.menu_item(u'&Edit->UnGroup\\tCtrl+U').click()


print3dItem(app, EzCadAppRef)