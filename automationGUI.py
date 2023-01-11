import tkinter
import xml.dom.minidom

from tkinter import filedialog


class App:
    def __init__(self, root):
        self.root = root
        self.root.minsize(640, 220)
        self.root.maxsize(640, 220)
        self.root.title("EzCad Automator")

        self.select_file_section = tkinter.LabelFrame(self.root, text="Import File")

        # window variables
        self.file_location = tkinter.StringVar()
        self.file_location.set("Enter location of file to be imported")
        self.no_of_layers = tkinter.IntVar()
        self.no_of_layers.set(0)
        self.layer_count_message = tkinter.StringVar()

        self.file_entry = tkinter.Entry(self.select_file_section, textvariable=self.file_location, width=50, relief="groove")
        self.file_entry.grid(row=0, column=0, padx=10, pady=10)

        self.browse_button = tkinter.Button(self.select_file_section, text="Browse...", command=self.browse_file, width=15, relief="groove")
        self.browse_button.grid(row=0, column=1, padx=20, pady=10)
        
        # self.no_of_layers_entry = tkinter.Label(self.select_file_section, text="Select a file to show its layer count", width=40)
        # self.no_of_layers_entry.config(textvariable=self.layer_count_message)
        # self.no_of_layers_entry.grid(row=1, padx=20, pady=10)

        self.start_button = tkinter.Button(self.select_file_section, text="Start", command=self.start_process, width=15, relief="groove")
        self.start_button.grid(row=1, column=0, pady=10)

        self.start_pause = tkinter.Button(self.select_file_section, text="Pause", command=self.pause_process, width=15, relief="groove")
        self.start_pause.grid(row=1, column=1, pady=10)

        self.start_resume = tkinter.Button(self.select_file_section, text="Resume", command=self.resume_process, width=15, relief="groove")
        self.start_resume.grid(row=1, column=2, pady=10, padx=10)

        self.select_file_section.grid(padx=10, pady=10)


    def browse_file(self):
        file_path = filedialog.askopenfilename(title="Select a File", filetypes=(("All Vector files", "*.ai *.plt *.dxf *.dst *.svg *.nc *.g *.gbr"), ("all files", "*.*")))

        if(file_path):
            self.file_location.set(file_path)
            self.calc_no_of_layers()


    def calc_no_of_layers(self):

        # parse the SVG file
        doc = xml.dom.minidom.parse(self.file_location.get())

        # count the number of 'g' elements (layers)
        layers = doc.getElementsByTagName('g')
        num_layers = len(layers)

        self.no_of_layers.set(num_layers)

        print(f'There are {num_layers} layers in this SVG file.')



    def start_process(self):
        pass


    def pause_process(self):
        pass


    def resume_process(self):
        pass


root = tkinter.Tk()
app = App(root)
root.mainloop()