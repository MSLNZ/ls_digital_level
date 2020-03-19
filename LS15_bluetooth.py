# https://stackoverflow.com/questions/16938647/python-code-for-serial-data-to-print-on-window

# Program works in Q-Level mode on LS15

# [x] parse str_in

# [ ] add some sounds for feedback around sucessful and unsucessful data
# [-] consider breaking if serial port is not open in SerialThread.run
# [-] clear buffer after read or failure to parse. reset_input_buffer()

# [x] add a count  field to display
# [x] on receiving string, query compass - note can set local magnetic declination
# [x] sometimes the height ID is 331, sometimes 332 ??!! backsight/foresight
# [x] put height to Excel current cell
# [x] convert distance, angle, height to x, y, z CSY
# [x] log raw strings to file (even use a logger)
# [x] log all data, height distance, angle, x, y, z to file
# [x] reformat, rename buttons
# [x] stringvar for the labels?
# [ ] install on other machines - cxfreeze or miniconda
# [x] drop down to select serial port
# [x] option to copy to clipboard, (x] interface
# [x] options for which data to go to excel or clipboard [x] interface
# [x] resolution of compass may only be 0.1 deg make display match
# [ ] interface asthetics - fixed sizes, borders to group settings
# [-] displays for x, y, z ?
# [x] visual confirmation workbook link is set
# [x] test bluetooth by reading connection and indicating its working
# [x] requires to be run as adminstrator can I bypass this for a simple bat file?

# [ ] not finding com ports on HP elitebook, win 10 problem?
#      same problem if 

from math import pi, cos, sin
import os
import datetime


import serial
import serial.tools.list_ports
import threading
import queue
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32

instrument_ids = ["700193", "701020"]


class SerialThread(threading.Thread):
    def __init__(self, queue, serial_port):
        threading.Thread.__init__(self)
        self.queue = queue
        self.sp = serial_port

    def run(self):
        while True:
            if self.sp.isOpen():
                if self.sp.inWaiting():
                    text = self.sp.readline(self.sp.inWaiting())
                    self.queue.put(text)


class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("LS15 Digital Level Reader")
        self.excel = None
        self.count = 0

        self.serial_port = serial.Serial()
        self.serial_port.baudrate = 115200
        # get list of serial ports
        ports = serial.tools.list_ports.comports()
        self.port_ids = sorted([port[0] for port in ports])
        if "COM4" in self.port_ids:
            self.serial_port.port = "COM4"
        else:
            self.serial_port.port = self.port_ids[0]

        self.create_display()

        self.queue = queue.Queue()
        thread = SerialThread(self.queue, self.serial_port)
        thread.start()
        self.mmt_count = 0

        # logging file names
        self.log_name = os.path.join(
            os.path.join(os.environ["USERPROFILE"]),
            "Desktop",
            "LS15_" + "{:%Y%m%d_%H%M}".format(datetime.datetime.now()) + ".log",
        )

        self.csv_name = os.path.join(
            os.path.join(os.environ["USERPROFILE"]),
            "Desktop",
            "LS15_" + "{:%Y%m%d_%H%M}".format(datetime.datetime.now()) + ".csv",
        )

    def process_serial(self):
        while self.queue.qsize():
            try:
                s = self.queue.get()
                log = open(self.log_name, "a")
                self.datestamp = datetime.datetime.now().isoformat(" ")
                log.write("{} {}".format(self.datestamp, str(s)))

                # parse string

                if s[25:27] == b"32" and (s[49:52] == b"331" or s[49:52] == b"332"):
                    # ask for compass reading
                    self.serial_port.write(b"%R1Q,29072:\r\n")
                    self.distance = float(s[32:48]) / 1e5
                    self.distance_var.set("{:.5f}".format(self.distance))
                    self.height = float(s[56:72]) / 1e5
                    self.height_var.set("{:.5f}".format(self.height))

                elif s[0:4] == b"%R1P":
                    # id string
                    if s[12:16] == b"LS15":
                        self.label_s0a["text"] = "Bluetooth connected"

                    else:
                        # compass string
                        self.angle = float(s[11:-2])
                        self.angle_var.set("{:.2f}".format(180.0 * self.angle / pi))
                        self.count += 1
                        self.count_var.set("{:>10d}".format(self.count))

                        # data outputs
                        if self.excel and self.data_excel != "None":
                            self.send_to_excel()
                        if self.data_clip != "None":
                            self.send_to_clipboard()

                        self.write_to_file()
                else:
                    log.write(" bad string")

                log.write("\r\n")
                log.close()

            except queue.Empty:
                pass
        self.after(100, self.process_serial)

    def calc_XYZ(self):
        self.X = self.distance * cos(self.angle)
        self.Y = self.distance * sin(self.angle)
        self.Z = self.height

    def send_to_excel(self):
        if self.move_excel.get() == "across":
            offset_A = (1, 2)
            offset_B = (2, 1)
            offset_C = (3, 1)
            offset_D = (4, 1)
            offset_E = (5, 1)
            offset_F = (6, 1)
            offset_G = (7, 1)
        else:
            offset_A = (2, 1)
            offset_B = (1, 2)
            offset_C = (1, 3)
            offset_D = (1, 4)
            offset_E = (1, 5)
            offset_F = (1, 6)
            offset_G = (1, 7)

        if self.data_excel.get() == "Height":
            self.excel.ActiveCell.Value = self.height_var.get()
            self.excel.ActiveCell.Offset(*offset_A).Select()

        if self.data_excel.get() == "Height, Distance, Bearing":
            self.excel.ActiveCell.Value = self.height_var.get()
            self.excel.ActiveCell.Offset(*offset_B).Value = self.distance_var.get()
            self.excel.ActiveCell.Offset(*offset_C).Value = self.angle_var.get()
            self.excel.ActiveCell.Offset(*offset_A).Select()

        if self.data_excel.get() == "X, Y, Z":
            self.calc_XYZ()
            self.excel.ActiveCell.Value = "{:.5f}".format(self.X)
            self.excel.ActiveCell.Offset(*offset_B).Value = "{:.5f}".format(self.Y)
            self.excel.ActiveCell.Offset(*offset_C).Value = "{:.5f}".format(self.Z)
            self.excel.ActiveCell.Offset(*offset_A).Select()

        if self.data_excel.get() == "All":
            self.calc_XYZ()
            self.excel.ActiveCell.Value = self.datestamp
            self.excel.ActiveCell.Offset(*offset_B) = self.height_var.get()
            self.excel.ActiveCell.Offset(*offset_C).Value = self.distance_var.get()
            self.excel.ActiveCell.Offset(*offset_D).Value = self.angle_var.get()
            self.excel.ActiveCell.Offset(*offset_E).Value = "{:.5f}".format(self.X)
            self.excel.ActiveCell.Offset(*offset_F).Value = "{:.5f}".format(self.Y)
            self.excel.ActiveCell.Offset(*offset_G).Value = "{:.5f}".format(self.Z)
            self.excel.ActiveCell.Offset(*offset_A).Select()

    def send_to_clipboard(self):
        """
        comma separated values
        """
        self.clipboard_clear()

        if self.data_clip.get() == "Height":
            self.clipboard_append(self.height_var.get())

        if self.data_clip.get() == "Height, Distance, Bearing":
            self.clipboard_append(
                "{}, {}, {}".format(
                    self.height_var.get(), self.distance_var.get(), self.angle_var.get()
                )
            )

        if self.data_clip.get() == "X, Y, Z":
            self.calc_XYZ()
            self.clipboard_append(
                "{:.5f}, {:.5f}, {:.5f}".format(self.X, self.Y, self.Z)
            )

        if self.data_clip.get() == "All":
            self.calc_XYZ()
            self.clipboard_append(
                "{}, {}, {}, {}, {:.5f}, {:.5f}, {:.5f} \r\n".format(
                self.datestamp,
                self.height_var.get(),
                self.distance_var.get(),
                self.angle_var.get(),
                self.X,
                self.Y,
                self.Z,
                )
            )

    def write_to_file(self):
        self.calc_XYZ()
        csv = open(self.csv_name, "a")
        csv.write(
            "{}, {}, {}, {}, {:.5f}, {:.5f}, {:.5f} \r\n".format(
                self.datestamp,
                self.height_var.get(),
                self.distance_var.get(),
                self.angle_var.get(),
                self.X,
                self.Y,
                self.Z,
            )
        )
        csv.close()

    def open_serial(self):
        print("opening port, expect bluetooth window on PC")
        self.serial_port.port = self.port_id.get()
        self.serial_port.open()
        # check connection by requesting LS15 id
        self.serial_port.write(b"%R1Q,5004:\r\n")
        self.process_serial()

    def close_serial(self):
        print("Goodbye")
        self.serial_port.close()

    def open_workbook(self):
        excel_file = filedialog.askopenfilename(title="Open Excel Workbook")
        if excel_file:
            self.excel = None
            self.excel = win32.gencache.EnsureDispatch("Excel.Application")
            wb = self.excel.Workbooks.Open(excel_file)
            self.excel.Visible = True
            self.label_e0["text"] = excel_file

    def create_display(self):

        # -----------------Serial Port column-----------------------------------
        self.open_serial_btn = tk.Button(
            self, text="Open Serial Port", command=self.open_serial
        )
        self.open_serial_btn.grid(row=10, column=0, sticky="W")

        self.label_s0a = tk.Label(self, text="No bluetooth ")
        self.label_s0a.grid(row=12, column=0, sticky="W")

        self.label_s0 = tk.Label(self, text="Com port")
        self.label_s0.grid(row=15, column=0, sticky="W")

        self.port_id = tk.StringVar(self)
        self.port_id.set(self.serial_port.port)
        self.port_drop = tk.OptionMenu(self, self.port_id, *self.port_ids)
        self.port_drop.grid(row=20, column=0, stick="W")

        self.close_serial_btn = tk.Button(
            self, text="Close Serial Port", command=self.close_serial
        )
        self.close_serial_btn.grid(row=25, column=0, sticky="W")

        # ----------------Excel column-----------------------------------

        self.open_workbook_btn = tk.Button(
            self, text="Open Excel Workbook", command=self.open_workbook
        )
        self.open_workbook_btn.grid(row=10, column=1, sticky="W")

        self.label_e0 = tk.Label(self, text="No excel connection")
        self.label_e0.grid(row=12, column=1, sticky="W")

        self.label_e1 = tk.Label(self, text="Data to Transfer")
        self.label_e1.grid(row=15, column=1, sticky="W")

        data_options = ["None", "Height", "Height, Distance, Bearing", "X, Y, Z", "All"]
        self.data_excel = tk.StringVar(self)
        self.data_excel.set("Height")
        self.data_excel_drop = tk.OptionMenu(self, self.data_excel, *data_options)
        self.data_excel_drop.grid(row=20, column=1, stick="W")

        self.label_e2 = tk.Label(self, text="Move Active Cell")
        self.label_e2.grid(row=25, column=1, sticky="W")

        move_options = ["across", "down"]
        self.move_excel = tk.StringVar(self)
        self.move_excel.set("across")
        self.move_excel_drop = tk.OptionMenu(self, self.move_excel, *move_options)
        self.move_excel_drop.grid(row=30, column=1, stick="W")

        # ----------------Clipboard column-----------------------------------
        self.label_c0 = tk.Label(self, text="Copy data to Clipboard")
        self.label_c0.grid(row=10, column=2, sticky="W")

        self.label_c1 = tk.Label(self, text="Data to Transfer")
        self.label_c1.grid(row=15, column=2, sticky="W")

        self.data_clip = tk.StringVar(self)
        self.data_clip.set("None")
        self.data_clip_drop = tk.OptionMenu(self, self.data_clip, *data_options)
        self.data_clip_drop.grid(row=20, column=2, stick="W")

        font1 = "Arial 36"

        # --------------Displays-----------------------------------------------

        self.label_c0 = tk.Label(self, text="Count", font=font1)
        self.label_c0.grid(row=50, column=0, sticky="W")

        self.count_var = tk.StringVar(self)
        self.count_var.set("----")
        self.label_count = tk.Label(self, textvar=self.count_var, font=font1)
        self.label_count.grid(row=50, column=1, sticky="E")

        # --------------------------------------------------------------------

        self.label2 = tk.Label(self, text="Height", font=font1)
        self.label2.grid(row=60, column=0, sticky="W")

        self.height_var = tk.StringVar(self)
        self.height_var.set("----.-----")
        self.label_height = tk.Label(self, textvar=self.height_var, font=font1)
        self.label_height.grid(row=60, column=1, sticky="W")

        self.label_m2 = tk.Label(self, text=" m", font=font1)
        self.label_m2.grid(row=60, column=2, sticky="W")

        # --------------------------------------------------------------------

        self.label3 = tk.Label(self, text="Distance   ", font=font1)
        self.label3.grid(row=80, column=0, sticky="W")

        self.distance_var = tk.StringVar(self)
        self.distance_var.set("----.-----")
        self.label_distance = tk.Label(self, textvar=self.distance_var, font=font1)
        self.label_distance.grid(row=80, column=1, sticky="W")

        self.label_m3 = tk.Label(self, text=" m", font=font1)
        self.label_m3.grid(row=80, column=2, sticky="W")

        # --------------------------------------------------------------------
        # use decimal degree

        self.label4 = tk.Label(self, text="Bearing", font=font1)
        self.label4.grid(row=100, column=0, sticky="W")

        self.angle_var = tk.StringVar(self)
        self.angle_var.set("----.-----")
        self.label_angle = tk.Label(self, textvar=self.angle_var, font=font1)
        self.label_angle.grid(row=100, column=1, sticky="W")

        self.label_deg = tk.Label(self, text=" °", font=font1)
        self.label_deg.grid(row=100, column=2, sticky="W")


app = App()
app.mainloop()
