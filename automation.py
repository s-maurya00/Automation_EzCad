import time, tkinter
import xml.dom.minidom
import threading

# import automationGUI_control
# from automationGUI_control import should_pause

from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from tkinter import filedialog



"""
INPUTS: 
    Path of SVG File,
    Number of Objects in SVG File,
    Time taken by Mark(F2) function (not 100% sure)
    Add PAUSE and RESUME features which are mandatory

FUTURE WORK:
    Remove the error functions,
    Take input for numberOfObjectsInSVG or calcuate it within the app,
    Calculate time required for Mark(F2) process to terminate,

    Create .exe file with appropriate inputs
"""


"""
# This is the section responsible for the pause and play buttons functioning
should_pause = False

class GUI_controls():

    def __init__(self, root):
        self.root = root
        self.root.minsize(350, 200)
        self.root.maxsize(350, 200)
        self.root.title("EzCad Automator controls")

        self.control_section = tkinter.LabelFrame(self.root, text="Controls")

        # Sets the window to always on top
        self.root.attributes("-topmost", True)


        # window variables
        # current_loop_count = 

        self.start_pause = tkinter.Button(self.control_section, text="Pause", command=self.pause_process, width=16, relief="groove")
        self.start_pause.grid(row=1, column=1, padx=10, pady=10)

        self.start_resume = tkinter.Button(self.control_section, text="Resume", command=self.resume_process, width=16, relief="groove")
        self.start_resume.grid(row=1, column=2, padx=10, pady=10)

        self.control_section.grid(row=0, padx=10, pady=10)

        self.layer_info_frame = tkinter.LabelFrame(self.root, text="Layer Information")

        self.layer_info = tkinter.Label(self.layer_info_frame, text=f"Printing layer: 12 of 275")
        self.layer_info.grid(row=2, padx=50, pady=10)

        self.layer_info_frame.grid(row=3, padx=10, pady=10)

    
    def pause_process(self):
        global should_pause
        should_pause = True

    def resume_process(self):
        global should_pause
        should_pause = False
"""






def selectFirstObjectInList(EzCadAppRef, loopCount):
    # Selects the first item to be printed from the available objects in the object list

    if(loopCount == 0):     # This reduces the PC Load by only having to click outside the Object List box ONCE
        # Select the "Edit Node" Tool
        EzCadAppRef[u'Toolbar3'].button(1).click()

        # Reselect the "Pick" Tool
        EzCadAppRef[u'Toolbar3'].button(0).click()

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


def clickEnableInHatching(EzCadAppRef, setCheckmarkTo):
    # Enable OR Disable the hatching checkbox in the "Hatching" Table and click "Apply"

    if(setCheckmarkTo == 0):   # '0' signifies that the checkbox has to be unchecked
        EzCadAppRef[u'Enable'].uncheck()
        EzCadAppRef[u'&Apply'].click_input()

    elif(setCheckmarkTo == 1): # '1' signifies that the checkbox has to be checkmarked
        EzCadAppRef[u'Enable'].check()
        EzCadAppRef[u'&Apply'].click_input()


def selectMarkingProperty(EzCadAppRef, iter):
    # Selects either 'black' or 'blue' marking property based on the 'iter' value

    if(iter == 0):
        EzCadAppRef[u'Button12'].click_input()
    elif(iter == 1):
        EzCadAppRef[u'Button13'].click_input()

    # Uncheck default param
    if(EzCadAppRef[u'Use default param'].get_check_state()):
        EzCadAppRef[u'Use default param'].click()

    # Set Speed(MM/Second)
    EzCadAppRef[u'Spin1'].click()
    speed = 10
    send_keys(f"{speed}")

    # Set Power%
    EzCadAppRef[u'Spin2'].click()
    power = 10
    send_keys(f"{power}")

    # Click Apply button
    EzCadAppRef[u'&Apply2'].click()


# INCOMPLETE
def startMarking(EzCadAppRef):
    # Start the Marking process

    EzCadAppRef[u'Mark(F2)'].click_input()
    time.sleep(2)   # Sleep till the printing is completed



def print3dItem(app, EzCadAppRef, numberOfObjectsInSVG, const_printing_interval):
    # Performs all relevant steps for all the available objects

    # global should_pause

    for i in range(0, numberOfObjectsInSVG):

        # while(should_pause):
        #     time.sleep(0.1)

        selectFirstObjectInList(EzCadAppRef, i)

        time.sleep(0.5)
        setXtoZero(EzCadAppRef)

        time.sleep(0.5)
        hatchObject(app, EzCadAppRef)

        time.sleep(0.5)
        clickEnableInHatching(EzCadAppRef, 1)   # Firstly, we enable the hatching for black one

        time.sleep(0.5)
        selectMarkingProperty(EzCadAppRef, 0)

        time.sleep(0.5)
        startMarking(EzCadAppRef)

        time.sleep(0.5)
        clickEnableInHatching(EzCadAppRef, 0)   # Here we disable the hatching for the blue one

        time.sleep(0.5)
        selectMarkingProperty(EzCadAppRef, 1)

        time.sleep(0.5)
        startMarking(EzCadAppRef)

        time.sleep(0.5)
        send_keys('{DELETE}')

        # Constant time that this function will pause between printing of each layer
        time.sleep(const_printing_interval)





def begin(const_printing_interval):

    app = Application().start(cmd_line=u'"C:\\Users\\maury\\Desktop\\For OptiLOM\\EZCAD2-Software\\EZCAD2 For AiO (20220915 Release)\\EZCAD2 For AiO.exe" ')
    EzCadAppRef = app[u'EzCad2.14.11 - No title']


    # Handling Errors
    app.License.IAgree.click()
    app.EzCad.Ok.click()
    app.EzCad.Ok.click()


    # Pause code's execuition untill the application gets loaded
    EzCadAppRef.wait('ready')


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



    """
    # Create the GUI window of controls on screen
    gui_controls_root = tkinter.Tk()
    gui_controls = GUI_controls(gui_controls_root)

    # gui_controls_root.after(1, print3dItem, app, EzCadAppRef, numberOfObjectsInSVG)  # This lets the print3dItem function to execuite even after the new gui is created

    gui_controls_root.mainloop()

    def print3d_thread():
        print3dItem(app, EzCadAppRef, numberOfObjectsInSVG)

    thread = threading.Thread(target=print3d_thread)
    thread.start()
    """
    time.sleep(1)
    print3dItem(app, EzCadAppRef, numberOfObjectsInSVG, const_printing_interval)



class GUI:
    def __init__(self, root):
        self.root = root
        self.root.minsize(620, 220)
        self.root.maxsize(620, 220)
        self.root.title("EzCad Automator")


        self.select_file_section = tkinter.LabelFrame(self.root, text="Import File")


        # window variables
        self.file_location = tkinter.StringVar()
        self.no_of_layers = tkinter.IntVar()
        self.no_of_layers.set(0)
        self.layer_count_message = tkinter.StringVar()
        self.const_printing_interval = tkinter.IntVar()
        self.const_printing_interval.set(0)


        # Define and positining of window elements
        self.file_entry = tkinter.Entry(self.select_file_section, textvariable=self.file_location, width=50, relief="groove")
        self.file_entry.grid(row=0, column=0, padx=10, pady=10)

        self.browse_button = tkinter.Button(self.select_file_section, text="Browse...", command=self.browse_file, width=15, relief="groove")
        self.browse_button.grid(row=0, column=1, padx=20, pady=10)


        # Take input for setting constant interval between printing of each layer
        self.const_printing_interval_entry = tkinter.Entry(self.select_file_section, textvariable=self.const_printing_interval, width=50, relief="groove")
        self.const_printing_interval_entry.grid(row=1, column=0, padx=10, pady=10)

        self.start_button = tkinter.Button(self.select_file_section, text="Start", command=self.start_process, width=15, relief="groove")
        self.start_button.grid(row=2, column=0, pady=10)


        self.select_file_section.grid(padx=10, pady=10)



    def browse_file(self):

        file_path = filedialog.askopenfilename(title="Select a File", filetypes=(("All Vector files", "*.ai *.plt *.dxf *.dst *.svg *.nc *.g *.gbr"), ("all files", "*.*")))

        if(file_path):
            self.file_location.set(file_path)



    def calc_no_of_layers(self):

        # parse the SVG file
        doc = xml.dom.minidom.parse(self.file_location.get())

        # count the number of 'g' elements (layers)
        layers = doc.getElementsByTagName('g')
        num_layers = len(layers)

        self.no_of_layers.set(num_layers)

        print(f'There are {num_layers} layers in this SVG file.')



    def start_process(self):

        if(self.file_location.get()):   # If file location is selected, only then should the program start
            gui_root.destroy()
            self.calc_no_of_layers()
            begin(self.const_printing_interval.get())


gui_root = tkinter.Tk()
gui = GUI(gui_root)
gui_root.mainloop()