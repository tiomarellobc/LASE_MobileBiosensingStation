from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as msg
from DMM import *
from Arduino import *
import pyvisa
import csv 
import random 
import time
import os
from threading import *
#Matplotlib

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg



window = Tk()
window.resizable(False,False)
secondary_window = Toplevel()
secondary_window.title("Live Plotting")

fig = Figure(figsize=(6,6), dpi=100)
LivePlot = fig.add_subplot(111)
LivePlot.set_xlabel("Gate Voltage (V)")
LivePlot.set_ylabel("Resistance (Ohms)")

SweepStatus = BooleanVar()
rm = pyvisa.ResourceManager()
print(rm.list_resources())

Keithley_2750 = None
myPort = None
Multimeter : DMM
DAC : Arduino

_filePath = ""
_fileName = ""


def Get_File_Path():
    _filePath = text=fd.askdirectory(initialdir='/')
    File_Path.config(text=_filePath)
    print(_filePath)

def Open_Ports():
    global Multimeter
    global DAC
    global Keithley_2750
    global myPort
    try:
        Keithley_2750 = rm.open_resource(Keithley_Address_Entry.get())
        Keithley_2750.timeout = 500000
        myPort = serial.Serial(Arduino_Address_Entry.get())
        DAC = Arduino(myPort, 0, 1500, 5)
        Multimeter = DMM(Keithley_2750)
        Multimeter.Reset()
        time.sleep(2)
        msg.showinfo("Port Status", f"Connection Successful : {Multimeter.Get_ID()} + {myPort.name}")
        window.title("Hardware Connected")
    except Exception:
        msg.showerror("Port Status", f"Connection Unsuccessful. Please ensure devices are connected with proper port names. Devices available are {rm.list_resources()}")
    
#Takes in user channel input and spits out proeprly formatted string
def Parse_Channels(input):
    All_Channels = []
    for Top_Channel in input.split(','):
        if(":" in Top_Channel):
            Bwoah = Top_Channel.split(':')
            low = int(Bwoah[0])
            high = int(Bwoah[1])+1
            for Sub_Channels in range(low, high):
                All_Channels.append(str(Sub_Channels))
        else:
            All_Channels.append(Top_Channel)
    return(All_Channels)


def Begin_Measurement_Thread():
    t1 = Thread(target=Begin_Measurement)
    t1.start()

def Begin_Measurement():
 
    global Multimeter
    global DAC
    global Keithley_2750
    global myPort
    try: 
        #Getting the channels that the user wants to scan; note
        Channel_selection = Channel_Selector.get().split(',')
        Channels = Parse_Channels(Channel_Selector.get())

    except Exception:
        msg.showerror("Channel Input", "Malformed Channel Information. Ensure a range of channels is provided. (Ex. 101:105)")
        return()

    #region DiracPointMeasurement
    if(SweepStatus.get()): #If we are sweeping the gate voltage
        start_vG = int(Vg_start_entry.get())
        end_vG = int(Vg_end_entry.get())
        delta_vG = int(Vg_delta_entry.get())
        mode = "a"
        
        DAC.Update_Gating(start_vG, end_vG, delta_vG)
        DAC.Set_Gate_Voltage(0)
        
        #Matplotlib Section
        LivePlot.clear()
        Plot_data = dict()
        Device_Data = dict()
        for Channel in Channels:
            Plot_data[Channel], = LivePlot.plot([], [], label=f"Device {Channel}")
            Device_Data[Channel] = ([],[])
        
        LivePlot.legend()
        LivePlot.set_xlim(0, end_vG/1000)
        ###End Matplotlib Section

        if(os.path.exists(f"Data/RawData/{File_Name.get()}.csv")):
            mode = "a"
        else:
            mode = "w"

        with open(f"Data/RawData/{File_Name.get()}.csv", mode, newline="") as file:
            writer = csv.writer(file, dialect='excel')
            #Creating the Header row of the csv
            Header = ["Gate Voltage (V)"]
            for channel in Channels:
                Header.append(channel)
            writer.writerow(Header)
            
            

            #Stepping through each Gate voltage and measuring resistance of the devices
            for Vg in DAC.Return_Gate_Voltages():
                window.title(f"Measuring Dirac Points | Current Gate Voltage:{Vg}")
                DAC.Set_Gate_Voltage(Vg)
                time.sleep(0.5)
                Resistances = Multimeter.Scan_Channels(",".join(Channels))
                #At this point, we have a list of resistance values for the devices, in order that they are supplied
                
                #Awful Awful, code
                
                for i in range(len(Resistances)):
                    #Updating the Matplotlib 
                    Device_Data[Channels[i]][0].append(Vg/1000)
                    Device_Data[Channels[i]][1].append(Resistances[i])
                    Plot_data[Channels[i]].set_xdata(Device_Data[Channels[i]][0])
                    Plot_data[Channels[i]].set_ydata(Device_Data[Channels[i]][1])
                    LivePlot.set_ylim(0, (max(Resistances))*1.50)
                    
                    print(f"Vg Data{Device_Data[Channels[i]]}")
                
                fig.canvas.draw()
                fig.canvas.flush_events()

                #End of Awful, Awful code
                Resistances.insert(0, Vg/1000)
                writer.writerow(Resistances)
                print(f"Current Gate Voltage: {Vg} Resistances: {Resistances}")
            
            
            
            DAC.Set_Gate_Voltage(0)
        msg.showinfo("Measurement Status", "Measurement Finished")
    #endregion
    
    #region Scanning Static Resistances
    else: #If we are not sweewping the gate voltage; just scan the resistnces
        DAC.Set_Gate_Voltage(0)
        mode = "a"
        if(os.path.exists(f"Data/RawData/{File_Name.get()}.csv")):
            mode = "a"
        else:
            mode = "w"

        window.title(f"Scanning Resistances of Channel {Start_Channel} to Channel {End_Channel}")
        with open(f"Data/RawData/{File_Name.get()}.csv", mode, newline="") as file:
            writer = csv.writer(file, dialect='excel')
            Header = []
            Res_Row = []
            for channel in range(Start_Channel, End_Channel+1, 1):
                Header.append(f"Device {channel}")
            Resistances = Multimeter.Scan_Channels(Start_Channel, End_Channel)
            print(Resistances)
            for val in Resistances:
                Res_Row.append(val)
            writer.writerow(Header)
            writer.writerow(Res_Row)
        msg.showinfo("Measurement Status", "Measurement Finished")
    #endregion


#region UI SETUP


window.title("HARDWARE NOT CONNECTED")

channel_label = Label(window, text="Channel Select 101:120")
Channel_Selector = Entry(window)

Vg_start = Label(window, text="Gate Voltage Start (mV)")
Vg_start_entry = Entry()
Vg_end = Label(window, text="Gate Voltage End (mV)")
Vg_end_entry = Entry()
Vg_delta = Label(window, text="Gate Voltage Delta (mV)")
Vg_delta_entry = Entry()

Sweep = Checkbutton(window, text="Sweep Gate Voltage?", onvalue=True, offvalue=False, variable=SweepStatus)
Set_File_Dialog = Button(window, text="Select File Location", command=Get_File_Path)
File_Path = Label(window, text="File Path Not Selected", width=20)
File_Name = Entry(window)
File_Name.insert(0, "Enter File Name!")
Measure = Button(window, text="Begin Measurement", command=Begin_Measurement_Thread)

Keithley_Address_Label = Label(window, text="Keithley Address")
Arduino_Address_Label = Label(window, text="Arduino Address")
Keithley_Address_Entry = Entry(window)
Keithley_Address_Entry.insert(0,"ASRL4::INSTR")
Arduino_Address_Entry = Entry(window)
Arduino_Address_Entry.insert(0,"COM3")
Open_Ports_Button = Button(window, text="Open Ports", command=Open_Ports)

Keithley_Address_Label.grid(row=0, column=0)
Keithley_Address_Entry.grid(row=0, column=1)
Arduino_Address_Label.grid(row=0, column=2)
Arduino_Address_Entry.grid(row=0, column=3)
Open_Ports_Button.grid(row=0, column=4)

channel_label.grid(row=1,column=0)
Channel_Selector.grid(row=1, column = 1)

Vg_start.grid(row=2, column=0)
Vg_start_entry.grid(row=2, column=1)
Vg_end.grid(row=2, column=2)
Vg_end_entry.grid(row=2, column=3)
Vg_delta.grid(row=2, column=4)
Vg_delta_entry.grid(row=2, column=5)
Sweep.grid(row=2, column=6)
Set_File_Dialog.grid(row=3, column=0)
File_Path.grid(row=3, column=2)
File_Name.grid(row=3, column=1)
Measure.grid(row=4,column=0)




x = [1,2,3,4,5]
y = [2,3,6,10,20]
z = [1,2,1,2,1]



canvas = FigureCanvasTkAgg(fig, master=secondary_window)
canvas.draw()
canvas.get_tk_widget().pack()
#endregion



window.mainloop()

