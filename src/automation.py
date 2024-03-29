# Date: 28-02-2023
import threading, time, os, datetime, tkinter, win32gui, xml.dom.minidom, logging, socket

import pyautogui

from pywinauto.application import Application
from pywinauto.keyboard import send_keys

from tkinter import filedialog, messagebox


# Basic config details of error logging file
logging.basicConfig(filename='errorlog.txt', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')


# PROGRAM CONSTANTS
WINDOW_FOCUS_WAIT_TIME = 0.1    # Time to wait for focusing of a window before clicking any button of EzCad window
ANY_CLICK_WAIT_TIME = 0.1       # Time to wait for any random click's task to be execuited by EzCad 
FUNCTION_INTERVAL_WAIT_TIME = 0.2   # Time to wait after performing a function in print3dItem()
RETRY_WAIT_TIME = 0.2   # Time to wait while retrying execuition of a function in print3dItem()


"""
INPUTS: 
    Path of SVG File,
    Number of Objects in SVG File,
    Time taken by Mark(F2) function (not 100% sure)
    Add PAUSE and RESUME features which are mandatory


    19-01-2023
    Take Marking delay input from the user in seconds
    

    21-01-2023(Tasks)
    To add play pause and resume functionality --> DONE
    To Implement Threading approach --> DONE
    To set marking wait time for window to disappear
    To declare the current_loop_count variable such that it is not above print3dItem function (it just looks ugly nothing else)
    To update the layer count dynamically in the 'Automation Controls' GUI window
    
    Break file into multiple small functional files and test its working


    25-01-2023(Tasks done)
    To update layer count in GUI dynamically
    To correct the calculate_number_of_layers button in the gui
    To resolve infinite threding problem
    To set mark function to wait till window disappears

    Layer count updation in automation control's GUI window
    Resolved problem of Creation of infinitely multiple threads in Control's resume function



FUTURE WORK:
    Remove the error functions,
    Take input for numberOfObjectsInSVG or calcuate it within the app,
    Calculate time required for Mark(F2) process to terminate,

    Create .exe file with appropriate inputs
"""





def deleteFirstNObjectsInList(EzCadAppRef, begin_at_layer_no):
    # Deletes the starting 'n' objects so that program starts at a given layer count

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    for i in range(0, begin_at_layer_no - 2):
        send_keys('+{DOWN}')

    send_keys('{DELETE}')



def selectFirstObjectInList(EzCadAppRef):
    # Selects the first item to be printed from the available objects in the object list

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    # if(loopCount == 0):     # This reduces the PC Load by only having to click outside the Object List box ONCE
    # Select the "Edit Node" Tool
    EzCadAppRef[u'Toolbar3'].button(1).click()

    # Reselect the "Pick" Tool
    EzCadAppRef[u'Toolbar3'].button(0).click()

    # Click the title of "Object List"
    objectHeader = EzCadAppRef[u'Header']
    objectHeader.click()
    # "Home" key selects the 1st element in the object list
    send_keys('{HOME}')



def showRedBoundary(app, EzCadAppRef):
    # Shows a red bounday around an item of the list which is used for affixing the base under the object

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    EzCadAppRef[u'Red(F1)'].click()

    while True:
        if(app[u'Mark'].exists()):   # If the Red boundary window is open, wait 1 second and recheck if it's open, if yes then wait again else return from the function
            time.sleep(1)
        else:
            break


def selectLastObjectInList(EzCadAppRef):
    # Selects the Last item from the object list of available objects as it is the currently hatched item

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    # Select the "Edit Node" Tool
    EzCadAppRef[u'Toolbar3'].button(1).click()

    # Reselect the "Pick" Tool
    EzCadAppRef[u'Toolbar3'].button(0).click()

    # Click the title of "Object List"
    objectHeader = EzCadAppRef[u'Header']
    objectHeader.click()
    # "End" key selects the Last element in the object list
    send_keys('{END}')



def setXtoZero(EzCadAppRef):
    # Sets the X axis position of current object to "0" and applies it

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    EzCadAppRef[u'Edit2'].type_keys('{END}+{HOME}0')
    EzCadAppRef[u'&Apply'].click_input()



def hatchObject(app, EzCadAppRef, is_programs_first_layer):
    # Hatches the current object and press 'OK'

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    # Open Hatching Menu window
    EzCadAppRef.menu_item(u'&Edit->Hatch\\tCtrl+H').click()

    time.sleep(ANY_CLICK_WAIT_TIME)
    if(is_programs_first_layer == 1): # Only wait for user input in setting the hatching property for the 1st loop item.
        while True:
            if app.Hatch.exists():  # If the Hatching property window is open, wait 1 second and recheck if it's open, if yes then wait again else return from the function
                time.sleep(1)
            else:
                break
    else:
        app.Hatch[u'&OK'].click_input()



def clickEnableInHatching(EzCadAppRef, setCheckmarkTo):
    # Enable OR Disable the hatching checkbox in the "Hatching" Table and click "Apply"

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    if(setCheckmarkTo == 0):   # '0' signifies that the checkbox has to be unchecked
        EzCadAppRef[u'Enable'].uncheck()
        EzCadAppRef[u'&Apply'].click_input()

    elif(setCheckmarkTo == 1): # '1' signifies that the checkbox has to be checkmarked
        EzCadAppRef[u'Enable'].check()
        EzCadAppRef[u'&Apply'].click_input()



def selectMarkingProperty(EzCadAppRef, iter, is_programs_first_layer):
    # Selects either 'black' or 'blue' marking property based on the 'iter' value

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    if(iter == 0):  # Black color parameters
        
        # Click Black Color button
        EzCadAppRef[u'Button12'].click_input()

        if(is_programs_first_layer == 1):

            # Uncheck default param
            if(EzCadAppRef[u'Use default param'].get_check_state()):
                EzCadAppRef[u'Use default param'].click()

            # Set Speed(MM/Second) for black parameter
            EzCadAppRef[u'Spin1'].click()
            speed = 750
            send_keys(f"{speed}")

            # Set Power% for black parameter
            EzCadAppRef[u'Spin2'].click()
            power = 50
            send_keys(f"{power}")

            # Click Apply button
            EzCadAppRef[u'&Apply2'].click()

    elif(iter == 1):

        # Click Blue color Button
        EzCadAppRef[u'Button13'].click_input()

        if(is_programs_first_layer == 1):

            # Uncheck default param
            if(EzCadAppRef[u'Use default param'].get_check_state()):
                EzCadAppRef[u'Use default param'].click()

            # Set Speed(MM/Second) for blue parameter
            EzCadAppRef[u'Spin1'].click()
            speed = 160
            send_keys(f"{speed}")

            # Set Power% for blue parameter
            EzCadAppRef[u'Spin2'].click()
            power = 70
            send_keys(f"{power}")

            # Click Apply button
            EzCadAppRef[u'&Apply2'].click()



def startMarking(app, EzCadAppRef, marking_time):
    # Start the Marking process

    # Set the focus to the EzCad software window
    time.sleep(WINDOW_FOCUS_WAIT_TIME)
    EzCadAppRef.set_focus()
    time.sleep(WINDOW_FOCUS_WAIT_TIME)

    EzCadAppRef[u'Mark(F2)'].click_input()
    # time.sleep(marking_time)   # Sleep till the printing is completed

    while True:
        if(app[u'Mark'].exists()):   # If the Marking property window is open, wait 1 second and recheck if it's open, if yes then wait again else return from the function
            time.sleep(1)
        else:
            break



current_loop_count = 0
retries = 0
is_programs_first_layer = 0     # This variable is used to know whether current layer is the 1st layer encountered by the program upon begining
def print3dItem(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time):

    """
        1. List of all called function in sequence is: 

            selectFirstObjectInList(EzCadAppRef)
            deleteFirstNObjectsInList(EzCadAppRef, begin_at_layer_no)

            selectFirstObjectInList(EzCadAppRef)
            showRedBoundary(app, EzCadAppRef)
            
            selectFirstObjectInList(EzCadAppRef)
            setXtoZero(EzCadAppRef)

            selectFirstObjectInList(EzCadAppRef)
            hatchObject(app, EzCadAppRef, current_loop_count)

            selectLastObjectInList(EzCadAppRef)
            clickEnableInHatching(EzCadAppRef, 1)

            selectLastObjectInList(EzCadAppRef)
            selectMarkingProperty(EzCadAppRef, 0, current_loop_count)

            selectLastObjectInList(EzCadAppRef)
            startMarking(app, EzCadAppRef, marking_time)

            selectLastObjectInList(EzCadAppRef)
            clickEnableInHatching(EzCadAppRef, 0)

            selectLastObjectInList(EzCadAppRef)
            selectMarkingProperty(EzCadAppRef, 1, current_loop_count)

            selectLastObjectInList(EzCadAppRef)
            startMarking(app, EzCadAppRef, marking_time)

            selectLastObjectInList(EzCadAppRef)
            send_keys('{DELETE}')
    """

    global should_pause, is_paused, current_loop_count, retries, is_programs_first_layer



    # Deletes the first "N" Objects and begins printing at layer "begin_at_layer_no"
    if(current_loop_count == 0 and begin_at_layer_no > 0):
        is_programs_first_layer = 1
        selectFirstObjectInList(EzCadAppRef)
        deleteFirstNObjectsInList(EzCadAppRef, begin_at_layer_no)
        current_loop_count += begin_at_layer_no - 1     # If begin_at_layer_no = 3 then current_loop_count must be = 2
        print(f"\nSuccessfully deleted first {begin_at_layer_no - 1} items!!")

        # Shows a red bounday around the 1st item of the list which is used for affixing the base under the object
        selectFirstObjectInList(EzCadAppRef)
        setXtoZero(EzCadAppRef)
        showRedBoundary(app, EzCadAppRef)
    

    # For myError_1; Added code for deleting just layer item cause it's nesting NG code is not generated automatically by "new_nest.cpp" file
    elif(current_loop_count == 0 and begin_at_layer_no <= 0):
        is_programs_first_layer = 1
        selectFirstObjectInList(EzCadAppRef)
        send_keys('{DELETE}')
        current_loop_count += 1
        print("\nSuccessfully deleted layer 1 item!!")

        # Shows a red bounday around the 1st item of the list which is used for affixing the base under the object
        selectFirstObjectInList(EzCadAppRef)
        setXtoZero(EzCadAppRef)
        showRedBoundary(app, EzCadAppRef)

    # Delete above section when NC code for layer 1 is genereated automatically

    # Setting the "is_programs_first_layer" to "0" so that rest of the program knows that current layer is not the 1st layer encountred by the program
    else:
        is_programs_first_layer = 0


    while((current_loop_count < numberOfObjectsInSVG) and (should_pause != True)):

        # Setting the "is_programs_first_layer" to "0" so that rest of the program knows that current layer is not the 1st layer encountred by the program
        # For myError_1; change the if condition to
        # "if(((begin_at_layer_no <= 0) and (current_loop_count > 0)) or ((begin_at_layer_no > 0) and (current_loop_count == begin_at_layer_no))):"
        # when the myError_1 is resolved
        if(((begin_at_layer_no <= 0) and (current_loop_count > 1)) or ((begin_at_layer_no > 0) and (current_loop_count == begin_at_layer_no))):
            is_programs_first_layer = 0


        # Setting X to Zero
        try:
            # Change to current_loop_count + 1 when myError_1 is resolved
            print("\nWorking on Layer: ", current_loop_count + 1)

            selectFirstObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Setting X coord to ZERO")
            setXtoZero(EzCadAppRef)

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in setXtoZero')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while setting X to Zero!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying setting X to Zero: ")

                    selectFirstObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Setting X coord to ZERO")
                    setXtoZero(EzCadAppRef)

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying setXtoZero in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in setXtoZero. Hence, Program is now paused.')
                should_pause = True
                break



        # Hatching the object
        try:
            selectFirstObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Hatching the object")
            hatchObject(app, EzCadAppRef, is_programs_first_layer)

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in hatchObject')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while Hatching Object!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to Hatch Object: ")

                    selectFirstObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Hatching the object")
                    hatchObject(app, EzCadAppRef, is_programs_first_layer)

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying hatchObject in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in hatchObject. Hence, Program is now paused.')
                should_pause = True
                break



        # Clicking the Enable Checkbox for Hatching
        try:
            selectLastObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Enabling Hatching")
            clickEnableInHatching(EzCadAppRef, 1)   # Firstly, we enable the hatching for black one

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in clickEnableInHatching(1)')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while Enabling Hatching!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to Click Enable in Hatching: ")

                    selectLastObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Enabling Hatching")
                    clickEnableInHatching(EzCadAppRef, 1)   # Firstly, we enable the hatching for black one

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying clickEnableInHatching(1) in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in clickEnableInHatching(1). Hence, Program is now paused.')
                should_pause = True
                break



        # Setting 1st Mark's properties
        try:
            selectLastObjectInList(EzCadAppRef) 

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Setting 1st Mark's Properties")
            selectMarkingProperty(EzCadAppRef, 0, is_programs_first_layer)

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in selectMarkingProperty(0)')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while setting 1st Mark's Property!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to set 1st Mark's Property: ")

                    selectLastObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Setting 1st Mark's Properties")
                    selectMarkingProperty(EzCadAppRef, 0, is_programs_first_layer)

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying selectMarkingProperty(0) in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in selectMarkingProperty(0). Hence, Program is now paused.')
                should_pause = True
                break



        # Starting the Lasor Marking
        try:
            selectLastObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Starting Marking")
            startMarking(app, EzCadAppRef, marking_time)

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in startMarking')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while starting Marking process!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to Start Marking: ")

                    selectLastObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Starting Marking")
                    startMarking(app, EzCadAppRef, marking_time)

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying startMarking in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in startMarking. Hence, Program is now paused.')
                should_pause = True
                break



        # Clicking the Disable Checkbox for Hatching
        try:
            selectLastObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Disabling Hatching")
            clickEnableInHatching(EzCadAppRef, 0)   # Here we disable the hatching for the blue one

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in clickEnableInHatching(0)')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while Disabling Hatching!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to Disable Hatching: ")

                    selectLastObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Disabling Hatching")
                    clickEnableInHatching(EzCadAppRef, 0)   # Here we disable the hatching for the blue one

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying clickEnableInHatching(0) in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in clickEnableInHatching(0). Hence, Program is now paused.')
                should_pause = True
                break



        # Selecting 2nd mark's property
        try:
            selectLastObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Setting 2nd Mark's Properties")
            selectMarkingProperty(EzCadAppRef, 1, is_programs_first_layer)

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in selectMarkingProperty(1)')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while setting 2nd Mark's Property!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to set 2nd Mark's Property: ")

                    selectLastObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Setting 2nd Mark's Properties")
                    selectMarkingProperty(EzCadAppRef, 1, is_programs_first_layer)

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying selectMarkingProperty(1) in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in selectMarkingProperty(1). Hence, Program is now paused.')
                should_pause = True
                break



        # Starting the Lasor Marking
        try:
            selectLastObjectInList(EzCadAppRef)

            time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
            print("\t- Starting Marking")
            startMarking(app, EzCadAppRef, marking_time)

        except Exception as e:
            logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in startMarking')

            # Retry the function call up to 3 times
            while retries < 3:
                print("\nSome unexpected error occured while starting Marking process!!!")
                print(f"\nRetry number: {retries + 1}")
                time.sleep(RETRY_WAIT_TIME)

                try:
                    print("\nRetrying to Start Marking: ")

                    selectLastObjectInList(EzCadAppRef)

                    time.sleep(FUNCTION_INTERVAL_WAIT_TIME)
                    print("\t- Starting Marking")
                    startMarking(app, EzCadAppRef, marking_time)

                    print(f"\nRetry SUCCESSFULL on retry number {retries + 1}!!!")
                    retries = 0
                    break

                except:
                    logging.exception(f'\n\n\nError number({retries + 1}), while retrying startMarking in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
                    print(f"\nRetry number {retries + 1} UNSUCCESSFULL!!!")

                retries += 1

            if(retries == 3):
                retries = 0
                logging.exception(f'\n\n\n MAX error retry count exceeded in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in startMarking. Hence, Program is now paused.')
                should_pause = True
                break



        # Set the focus to the EzCad software window
        time.sleep(WINDOW_FOCUS_WAIT_TIME)
        EzCadAppRef.set_focus()
        time.sleep(WINDOW_FOCUS_WAIT_TIME)

        print("\t- Deleting current layer item")
        selectLastObjectInList(EzCadAppRef)
        send_keys('{DELETE}')

        # except Exception as e:
        #     logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')

        #     # Retry the function call up to 3 times
        #     while retries < 3:
        #         print("\nSome unexpected error occured while running automation!!!")
        #         retries += 1
        #         print(f"\nRetry number: {retries}")
        #         time.sleep(1)
        #         try:
        #             # Change to current_loop_count + 1   when myError_1 is resolved
        #             print("\nRetrying Layer: ", current_loop_count + 2)

        #             selectFirstObjectInList(EzCadAppRef)

        #             time.sleep(0.5)
        #             print("\t- Setting X coord to ZERO")
        #             setXtoZero(EzCadAppRef)

        #             time.sleep(0.5)
        #             print("\t- Hatching the object")
        #             hatchObject(app, EzCadAppRef, current_loop_count)

        #             time.sleep(0.5)
        #             print("\t- Enabling Hatching")
        #             clickEnableInHatching(EzCadAppRef, 1)   # Firstly, we enable the hatching for black one

        #             time.sleep(0.5)
        #             print("\t- Setting 1st Mark's Properties")
        #             selectMarkingProperty(EzCadAppRef, 0, current_loop_count)

        #             time.sleep(0.5)
        #             print("\t- Starting Marking")
        #             startMarking(EzCadAppRef, marking_time)

        #             time.sleep(0.5)
        #             print("\t- Disabling Hatching")
        #             clickEnableInHatching(EzCadAppRef, 0)   # Here we disable the hatching for the blue one

        #             time.sleep(0.5)
        #             print("\t- Setting 2nd Mark's Properties")
        #             selectMarkingProperty(EzCadAppRef, 1, current_loop_count)

        #             time.sleep(0.5)
        #             print("\t- Starting Marking")
        #             startMarking(EzCadAppRef, marking_time)

        #             # Set the focus to the EzCad software window
        #             time.sleep(0.1)
        #             EzCadAppRef.set_focus()
        #             time.sleep(0.1)

        #             print("\t- Deleting current layer item")
        #             selectLastObjectInList(EzCadAppRef)
        #             send_keys('{DELETE}')

        #             print(f"\nRetry SUCCESSFULL on retry number {retries}!!!")
        #             retries = 0
        #             break
        #         except:
        #             logging.exception(f'\n\n\nRetrying Error number: {retries} in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()}')
        #             print(f"\nRetry number {retries} UNSUCCESSFULL!!!")

        if(should_pause != True):   # To check if pause has been pressed while printing, if 'yes', then don't sleep for new layer printing
            # Constant time that this function will pause between printing of each layer
            print(f"\t- Waiting for 'Next Layer Delay: {const_printing_interval} sec' time")

            # sleeping for 1 second as i want user to be able to pause execuition while nextLayerDelay is in progress
            for i in range(0, const_printing_interval):
                # Checking if pause button has been pressed while the program was waiting for Next Layer Delay Time, else sleep for 1 seconds for const_printing_interval time i.e., 1 * const_printing_interval = const_printing_interval time sleep
                if(should_pause == True):
                    is_paused = True
                    break
                time.sleep(1)
            # time.sleep(const_printing_interval)
    
        current_loop_count += 1


    if(should_pause == True):
        print("\nPaused")
        is_paused = True
    elif(should_pause == False):
        print("\nResumed")
        is_paused = False




def moveAppToRightSide():
    
    # successStatus = False
    pyautogui.FAILSAFE = False
    
    # Find the Notepad window
    hwnd = win32gui.FindWindow(None, "EzCad2.14.11 - No title")
    if hwnd:
        # Get the position of the title bar
        rect = win32gui.GetWindowRect(hwnd)
        x = rect[0] + rect[2] / 2
        y = rect[1] + 30
        # Move the cursor to the title bar and double-click
        pyautogui.moveTo(x, y, duration=0.5)
        # pyautogui.doubleClick()
        # Move the cursor to the top-right corner of the screen
        screen_width, screen_height = pyautogui.size()
        # pyautogui.moveTo(screen_width, 0, duration=0.5)
        pyautogui.dragTo(screen_width, screen_height/2, duration=0.5)
        # Release the click
        pyautogui.mouseUp()
        # successStatus = True

    else:
        messagebox.showerror("Error", "Unable to Auto-Reposition the window.")
        print("\nMyError: Window named \"EzCad2.14.11\" not was found.")
        # successStatus = False

    # return successStatus




def waitForResizing(resize_wait_time):

    def update_remaining_time():
        nonlocal resize_wait_time
        resize_wait_time -= 1

        label.config(text=f"Remaining Wait Time: {resize_wait_time} seconds")

        if resize_wait_time == 0:
            wait_window.destroy()
            time.sleep(1)
            return
        label.after(1000, update_remaining_time)

    wait_window = tkinter.Tk()
    wait_window.title("Time for arranging windows")

    try:
        wait_window.wm_iconbitmap("../assets/automated.ico")
    except:
        pass


    # Parent used for setting label vertically centered
    parent = tkinter.Frame(wait_window)

    wait_window.geometry("400x200")

    # Sets the window to always on top
    wait_window.attributes("-topmost", True)

    label = tkinter.Label(parent, text=f"Remaining Wait Time: {resize_wait_time} seconds")
    label.pack(fill='x')

    resizing_wait_time_label = tkinter.Label(parent, text="Manually select \"View Workspace\" button")
    resizing_wait_time_label.pack(padx=10, pady=10)

    label.after(1000, update_remaining_time)
    parent.pack(expand=1)

    wait_window.mainloop()





def play_process(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time):
    global should_pause, is_paused, thread

    # should_pause is used to tell the print3dItem function that the pause button has been pressed and the function has to be stopped after execution of current cycle which may take some "x" seconds
    # is_paused is used to know whether the "x" seconds required for the print3dItem function has been completed and the function has truely stopped
    should_pause = False
    is_paused = False

    print("\nStarting the Printing Process")

    thread = threading.Thread(target=print3dItem, args=(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time))
    thread.start()



def pause_process():
    global should_pause, is_paused
    # should_pause = True

    if(is_paused == False):
        if(should_pause == False):
            print("\nPause Pressed (Waiting for current loop to complete it's execution)\n")
            should_pause = True
        else:
            print("\nWaiting for current loop to complete execution\n")
    else:
        print("\nProcess is already PAUSED")
        # return

    # print("Waiting for current loop to complete execution\n")



def resume_process(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time, layer_info, gui_controls_root):
    global should_pause, thread, is_paused
    should_pause = False

    if(is_paused):
        print("\nResume Pressed")
    else:
        print("\nProcess is already RUNNING")
        return

    time.sleep(0.1) # Just for Safety purpose as parallel execution of process may give rise to errors
    if(is_paused):
        is_paused = False
        
        thread = threading.Thread(target=print3dItem, args=(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time))
        thread.start()

        update_current_layer_count(layer_info, numberOfObjectsInSVG, gui_controls_root)




def update_current_layer_count(layer_info, numberOfObjectsInSVG, gui_controls_root):
    global current_loop_count, should_pause

    layer_info.config(text=f"Printing layer: {current_loop_count + 1} of {numberOfObjectsInSVG}")

    if(should_pause == False):
        gui_controls_root.after(1000, lambda: update_current_layer_count(layer_info, numberOfObjectsInSVG, gui_controls_root))



def createControlGUI(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time):
    # Create the GUI window of controls on screen
    gui_controls_root = tkinter.Tk()
    
    gui_controls_root.title("Automator controls")

    try:
        gui_controls_root.wm_iconbitmap("../assets/automated.ico")
    except:
        pass

    gui_controls_root.minsize(305, 170)
    gui_controls_root.maxsize(305, 170)

    gui_controls_root.attributes("-topmost", True)

    gui_control_frame = tkinter.LabelFrame(gui_controls_root, text="Controls")


    play_process(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time)


    pause_button = tkinter.Button(gui_control_frame, text="Pause", command=pause_process, width=16, relief="groove")
    pause_button.grid(row=1, column=1, padx=10, pady=10)

    resume_button = tkinter.Button(gui_control_frame, text="Resume", command=lambda: resume_process(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time, layer_info, gui_controls_root), width=16, relief="groove")
    resume_button.grid(row=1, column=2, padx=10, pady=10)

    gui_control_frame.grid(row=0, padx=10, pady=10)


    layer_info_frame = tkinter.LabelFrame(gui_controls_root, text="Layer Information")

    layer_info = tkinter.Label(layer_info_frame, text=f"Printing layer: 1 of {numberOfObjectsInSVG}")
    layer_info.grid(row=2, padx=50, pady=10)
    update_current_layer_count(layer_info, numberOfObjectsInSVG, gui_controls_root)

    layer_info_frame.grid(row=3, padx=10, pady=10)


    gui_controls_root.mainloop()


    def print3d_thread(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time):
        print3dItem(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time)

    thread = threading.Thread(target=print3d_thread, args=(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time))
    thread.start()





def begin(resizingWaitTime, begin_at_layer_no, const_printing_interval, marking_time):

    app = Application(backend='win32').start(cmd_line=u'"C:\\Users\\maury\\Desktop\\For OptiLOM\\EZCAD2-Software\\EZCAD2 For AiO (20220915 Release)\\EZCAD2 For AiO.exe" ')
    EzCadAppRef = app[u'EzCad2.14.11 - No title']


    # Handling Errors
    if app.License.exists():
        app.License.IAgree.click()

    if app.EzCad.exists():
        app.EzCad.Ok.click()

    if app.EzCad.exists():
        app.EzCad.Ok.click()


    # Pause code's execuition untill the application gets loaded
    try:
        time.sleep(0.5)
        EzCadAppRef.wait('ready')

    except Exception as e:
        logging.exception(f'\n\n\nError in file {os.path.basename(__file__)} on {socket.gethostname()} at {datetime.datetime.now()} in begin function\'s EzCadAppRef.wait(\'ready\')')
        print("\nInitializing program, please wait while we set everything up for you.")
        time.sleep(5)

    # Double click titleBar of the window and drag it to right-corner to snap to the same
    time.sleep(2)
    moveAppToRightSide()
    time.sleep(2)


    # Wait for user to resize other windows
    waitForResizing(resizingWaitTime)


    # Open the window to import file
    EzCadAppRef.menu_item(u'&File->Import Vector File...\\tCtrl+B').click()


    # Type the file_location in input-box
    location = gui.file_location.get().replace("/", "\\")
    app.Open.Edit.type_keys(str(location), with_spaces = True)


    # Click 'OK' button in "Open" window
    time.sleep(0.5)
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


    # Either Calculate the total number of Objects or take the input from the user
    numberOfObjectsInSVG = gui.no_of_layers.get()


    createControlGUI(app, EzCadAppRef, numberOfObjectsInSVG, begin_at_layer_no, const_printing_interval, marking_time)





class GUI:
    def __init__(self, root):
        self.root = root
        self.root.minsize(700, 500)
        self.root.maxsize(700, 500)
        self.root.title("EzCad Automator")

        try:
            self.root.wm_iconbitmap("../assets/automated.ico")
        except:
            pass



        # window variables
        self.file_location = tkinter.StringVar()
        self.no_of_layers = tkinter.IntVar()
        self.no_of_layers.set(-1)
        self.layer_count_message = tkinter.StringVar()
        self.const_printing_interval = tkinter.IntVar()
        self.const_printing_interval.set(40)
        self.resizing_wait_time = tkinter.IntVar()
        self.resizing_wait_time.set(45)
        self.marking_time = tkinter.IntVar()
        self.marking_time.set(10)
        self.begin_at_layer_no = tkinter.IntVar()
        self.begin_at_layer_no.set(0)


        self.select_file_section = tkinter.LabelFrame(self.root, text="Import File")


        # Browse and input file url to be imported in Ezcad
        self.file_entry = tkinter.Entry(self.select_file_section, textvariable=self.file_location, width=50, relief="groove")
        self.file_entry.grid(row=0, column=0, padx=10, pady=10)

        self.browse_button = tkinter.Button(self.select_file_section, text="Browse...", command=self.browse_file, width=15, relief="groove")
        self.browse_button.grid(row=0, column=1, padx=20, pady=10)


        # Take input for setting constant interval between printing of each layer
        self.const_printing_interval_label = tkinter.Label(self.select_file_section, text="Next Layer Delay")
        self.const_printing_interval_label.grid(row=1, column=0, padx=10, pady=10)

        self.const_printing_interval_entry = tkinter.Entry(self.select_file_section, textvariable=self.const_printing_interval, width=15, relief="groove")
        self.const_printing_interval_entry.grid(row=1, column=1, padx=10, pady=10)


        # Take input for time to wait for user to resize their windows as they see fit
        self.resizing_wait_time_label = tkinter.Label(self.select_file_section, text="Window Resizing time")
        self.resizing_wait_time_label.grid(row=2, column=0, padx=10, pady=10)

        self.resizing_wait_time_entry = tkinter.Entry(self.select_file_section, textvariable=self.resizing_wait_time, width=15, relief="groove")
        self.resizing_wait_time_entry.grid(row=2, column=1, padx=10, pady=10)


        self.select_file_section.grid(padx=10, pady=10)


        # Take input for time to wait for marking to complete
        self.marking_time_label = tkinter.Label(self.select_file_section, text="Lasor movement time for hatching (not in use)")
        self.marking_time_label.grid(row=3, column=0, padx=10, pady=10)

        self.marking_time_entry = tkinter.Entry(self.select_file_section, textvariable=self.marking_time, width=15, relief="groove")
        self.marking_time_entry.grid(row=3, column=1, padx=10, pady=10)


        self.select_file_section.grid(padx=10, pady=10)




        self.calculate_layers_section = tkinter.LabelFrame(self.root, text="Click to calculate number of layers")


        # Calculate the total number of layers in imported file
        self.start_button = tkinter.Button(self.calculate_layers_section, text="Calculate number of Layers", command=self.calc_no_of_layers, width=30, relief="groove")
        self.start_button.grid(row=0, column=0, padx=10, pady=10)

        global calculated_no_of_layers_label
        calculated_no_of_layers_label = tkinter.Label(self.calculate_layers_section, text=f"Ans: {self.no_of_layers.get()} layers")
        calculated_no_of_layers_label.grid(row=0, column=1, padx=10, pady=10)
        calculated_no_of_layers_label.grid_remove()


        # Start at particular layer number
        self.begin_at_layer_no_label = tkinter.Label(self.calculate_layers_section, text="Begin at Layer number: ")
        self.begin_at_layer_no_label.grid(row=1, column=0, padx=10, pady=10)

        self.begin_at_layer_no_entry = tkinter.Entry(self.calculate_layers_section, textvariable=self.begin_at_layer_no, width=15, relief="groove")
        self.begin_at_layer_no_entry.grid(row=1, column=1, padx=10, pady=10)

        self.start_button = tkinter.Button(self.calculate_layers_section, text="Start", command=self.start_process, width=15, relief="groove")
        self.start_button.grid(row=2, column=0, pady=10, columnspan=2)

        self.calculate_layers_section.grid(padx=10, pady=10)
        


    def browse_file(self):

        file_path = filedialog.askopenfilename(title="Select a File", filetypes=(("All Vector files", "*.ai *.plt *.dxf *.dst *.svg *.nc *.g *.gbr"), ("all files", "*.*")))

        if(file_path):
            self.file_location.set(file_path)



    def calc_no_of_layers(self):

        # parse the SVG file
        if (self.file_location.get() != ''):
            doc = xml.dom.minidom.parse(self.file_location.get())

            # count the number of 'g' elements (layers)
            layers = doc.getElementsByTagName('g')
            num_layers = len(layers)

            self.no_of_layers.set(num_layers)

            # Set global 'resizing_wait_time_label' to display the calculate number of layers
            global calculated_no_of_layers_label
            calculated_no_of_layers_label.config(text=f"{self.no_of_layers.get()} layers")
            calculated_no_of_layers_label.grid()
            


    def start_process(self):

        if(self.file_location.get()):   # If file location is selected, only then should the program start
            gui_root.destroy()
            if(self.no_of_layers.get() <= 0):
                self.calc_no_of_layers()
            begin(self.resizing_wait_time.get(), self.begin_at_layer_no.get(), self.const_printing_interval.get(), self.marking_time.get())


gui_root = tkinter.Tk()
gui = GUI(gui_root)
gui_root.mainloop()