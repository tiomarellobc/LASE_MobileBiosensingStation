from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as msg
from DMM import *
from Arduino import *
import pyvisa
import csv 
import time
import os
from threading import *
import pandas as pd
#Matplotlib

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
window = Tk()
window.resizable(False,False)
secondary_window = Toplevel()
secondary_window.title("Live Plotting")

fig = Figure(figsize=(7,7), dpi=100)
LivePlot = fig.add_subplot(111)
LivePlot.set_xlabel("Gate Voltage (V)")
LivePlot.set_ylabel("Resistance (Ohms)")
LivePlot.set_xlim(0, 1.5)

SweepStatus = BooleanVar()
rm = pyvisa.ResourceManager()
print(rm.list_resources())
Keithley_2750 = None
myPort = None
Multimeter : DMM
DAC : Arduino
Aborted = False

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
        time.sleep(1)
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
    global Aborted

    try: 
        
        Channel_selection = Channel_Selector.get().split(',')
        Channels = Parse_Channels(Channel_Selector.get())

    except Exception:
        msg.showerror("Channel Input", "Malformed Channel Information. Ensure a range of channels is provided. (Ex. 101:105)")
        return()

    if(File_Path.cget("text") == "File Path Not Selected"):
        msg.showerror("File Path not selected. Please select, and try again.")
        return()

    
    start_vG = int(Vg_start_entry.get())
    end_vG = int(Vg_end_entry.get())
    delta_vG = int(Vg_delta_entry.get())
    mode = "a"
    excel_mode = "w"

    DAC.Update_Gating(start_vG, end_vG, delta_vG)
    DAC.Set_Gate_Voltage(0)

    if(Vg_SweepBack.getvar() == True):
        GateVoltages = DAC.Return_Gate_Voltages_SweepBack()
    else:
        GateVoltages = DAC.Return_Gate_Voltages()
    
    
    #Matplotlib Section
    LivePlot.clear()
    Plot_data = dict()
    Device_Data = dict()
    for Channel in Channels:
        Plot_data[Channel], = LivePlot.plot([], [], label=f"Device {Channel}")
        Device_Data[Channel] = ([],[])
    
    LivePlot.legend()
    LivePlot.set_xlim(0, end_vG/1000)
    LivePlot.set_xlabel("Gate Voltage (V)")
    LivePlot.set_ylabel("Resistance (Ohms)")
    ###End Matplotlib Section

    Recorded_Data = pd.DataFrame()

    path = File_Path.cget("text")
    final_path = path+"/"+File_Name.get()

    if(os.path.exists(final_path+".xlsx")):
        excel_mode = "a"
        print('previous file found')
    else:
        excel_mode = "w"

    #Stepping through each Gate voltage and measuring resistance of the devices
    for Vg in GateVoltages:
        if(Aborted):
            msg.showinfo("Aborted Measurement!", "Ending on next scan!")
            Aborted = False
            return()
        
        window.title(f"Measuring Dirac Points | Current Gate Voltage:{Vg}")
        DAC.Set_Gate_Voltage(Vg)
        time.sleep(2)
        Resistances = Multimeter.Scan_Channels(",".join(Channels))
        #At this point, we have a list of resistance values for the devices, in order that they are supplied
        
        #Awful Awful, code; who fucking wrote this shit? Oh wait, I did
        #I should never be allowed to write code; I'll have to streamline this later
        for i in range(len(Resistances)):
            #Updating the Matplotlib 
            Device_Data[Channels[i]][0].append(Vg/1000)
            Device_Data[Channels[i]][1].append(Resistances[i])
            Plot_data[Channels[i]].set_xdata(Device_Data[Channels[i]][0])
            Plot_data[Channels[i]].set_ydata(Device_Data[Channels[i]][1])
            LivePlot.set_ylim(0, (max(Resistances))*1.50)
            
        fig.canvas.draw()
        fig.canvas.flush_events()

        #End of Awful, Awful code
        Resistances.insert(0, Vg/1000)
        print(f"Current Gate Voltage: {Vg} Resistances: {Resistances}")
    
    DAC.Set_Gate_Voltage(0)

    #Inserting the gate voltage leftmost column
    Recorded_Data.insert(0, "Gate Voltages", Device_Data[Channels[0]][0])
    #Placing each column of resistance values, for each device, into the dataframe
    for i in range(len(Channels)):
        Channel_Header = Channels[i]+Channel_header_suffix_entry.get() #Adds the typed suffix
        Channel = Channels[i]
        Resistances = Device_Data[Channel][1]
        Recorded_Data.insert(len(Recorded_Data.columns), Channel_Header, Resistances)

    #Section places a row at the bottom of the dataframe, including the gate voltages at which the dirac point is observed
    #At this point, we have a dataframe loaded with a gatevoltage column, follwoed by columns of the ressitances values of the devices
    Max_VG = []
    for column_name, column_series in Recorded_Data.items():
        if(column_name == 'Gate Voltages'):
            Max_VG.append('Dirac Point Vg')
        else:
            Max_Resistance = column_series.max()
            index_of_max = Recorded_Data[Recorded_Data[column_name] == Max_Resistance].index.tolist()[0] #Dangerous, assumes that there is only one peak in the data, which may not be true
            Vg_at_peak = Recorded_Data['Gate Voltages'].iloc[index_of_max].tolist()
            Max_VG.append(Vg_at_peak)
    Recorded_Data.loc[len(Recorded_Data.index)] = Max_VG



    if(excel_mode == "a"):
        existing_file = pd.read_excel(final_path+".xlsx")
        existing_file.insert(len(existing_file.columns), "","")
        Final_Output = pd.concat([existing_file, Recorded_Data], axis=1)
        Final_Output.to_excel(final_path+'.xlsx', index=False)
    else:
        Recorded_Data.to_excel(final_path+'.xlsx', index=False)
        
    
    fig.savefig(final_path + File_Name.get() + Channel_header_suffix_entry.get() + ".png")


    msg.showinfo("Measurement Status", "Measurement Finished")

    
def Abort_Measurement():
    global Aborted
    Aborted = True

#region UI SETUP


window.title("HARDWARE NOT CONNECTED")

channel_label = Label(window, text="Channel Select 101:120")
Channel_Selector = Entry(window)
Channel_header_suffix_label = Label(window, text="Device Header Suffix")
Channel_header_suffix_entry = Entry(window)

Vg_start = Label(window, text="Gate Voltage Start (mV)")
Vg_start_entry = Entry()
Vg_end = Label(window, text="Gate Voltage End (mV)")
Vg_end_entry = Entry()
Vg_delta = Label(window, text="Gate Voltage Delta (mV)")
Vg_delta_entry = Entry()
Vg_SweepBack = Checkbutton()
Vg_SweepBack_Label = Label(window, text="SweepBack?")

Set_File_Dialog = Button(window, text="Select File Location", command=Get_File_Path)
File_Path = Label(window, text="File Path Not Selected", width=20)
File_Name = Entry(window)
File_Name.insert(0, "Enter File Name!")
Measure = Button(window, text="Begin Measurement", command=Begin_Measurement_Thread)
Abort_Measure = Button(window, text="Abort Measurement", command=Abort_Measurement)
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
Channel_header_suffix_label.grid(row=1, column=2)
Channel_header_suffix_entry.grid(row=1, column=3)

Vg_start.grid(row=2, column=0)
Vg_start_entry.grid(row=2, column=1)
Vg_end.grid(row=2, column=2)
Vg_end_entry.grid(row=2, column=3)
Vg_delta.grid(row=2, column=4)
Vg_delta_entry.grid(row=2, column=5)
Vg_SweepBack_Label.grid(row=2,column=6)
Vg_SweepBack.grid(row=2, column=7)
Set_File_Dialog.grid(row=3, column=0)
File_Path.grid(row=3, column=2)
File_Name.grid(row=3, column=1)
Measure.grid(row=4,column=0)
Abort_Measure.grid(row=4,column=1)



canvas = FigureCanvasTkAgg(fig, master=secondary_window)
canvas.draw()
canvas.get_tk_widget().pack()
#endregion



window.mainloop()

