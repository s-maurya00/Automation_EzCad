import tkinter

should_pause = False

# GUI window created on 10th Jan 2023 for controls like pause and resume
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


# gui_controls_root = tkinter.Tk()
# gui_controls = GUI_controls(gui_controls_root)
# gui_controls_root.mainloop()