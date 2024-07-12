from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as msg
from DMM import *
from Arduino import *
from GFET import *
import pyvisa
import csv 
import time
import os
from threading import *
import pandas as pd

class ThreadReturn(Thread):
    '''
    This thread is capable of returning a functions output
    '''
    def __init__(self, group=None, target=None, name=None,
                 args=(), kwargs={}, Verbose=None):
        Thread.__init__(self, group, target, name, args, kwargs)
        self._return = None

    def run(self):
        if self._target is not None:
            self._return = self._target(*self._args, **self._kwargs)
    
    def join(self, *args):
        Thread.join(self, *args)
        return self._return

#assigning core variables
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
window = Tk()
window.resizable(False,False)
secondary_window = Toplevel()
secondary_window.title("Live Plotting")
log_window = Toplevel()
log_window.title("Log")
log_window.geometry('650x400+50+300')

fig = Figure(figsize=(7,7), dpi=100)
LivePlot = fig.add_subplot(111)
LivePlot.set_xlabel("Gate Voltage (V)")
LivePlot.set_ylabel("Resistance (Ohms)")
LivePlot.set_xlim(0, 1.5)

rm = pyvisa.ResourceManager()
print(rm.list_resources())
Keithley_2750 = None
myPort = None
Multimeter : DMM
DAC : Arduino
Aborted = False

_filePath = ""
_fileName = ""

GFETS = []

k = 1
for i in range(1,21):
    device = f"W{k}D{i}"
    GFETS.append(GFET(100+i,device))
    if i % 5 == 0:
        k += 1

#region FUNCTIONS
def Get_File_Path():
    _filePath = text=fd.askdirectory(initialdir='/')
    file_path.config(text=_filePath)
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
    
def Parse_Channels(input):
    '''
    Takes in user channel input and returns properly formatted string
    Example: 101, 103, 104:106 ==> ['101','103','104','105','106']
    '''
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
    All_Channels = sorted(All_Channels)
    return(All_Channels)

def Begin_Manual_Measurement_Thread():
    '''
    This will initiate Begin_Measurement, pulling from the input values 
    '''
    channels = channel_entry.get()
    start_vg = int(start_vg_entry.get())
    end_vg = int(end_vg_entry.get())
    delta_vg = int(delta_vg_entry.get())

    t1 = ThreadReturn(target=Begin_Measurement,args=(channels,start_vg,end_vg,delta_vg))
    t1.start()

def Begin_Measurement(Channel_Input, Start_Vg, End_Vg, Delta_Vg):
    '''
    Takes 4 input, returns a dictionary of the dirac points in the format:
    dirac_dict = {'W#D##': [dirac], 'W#D##': [dirac], ...}
    '''
    global Multimeter
    global DAC
    global Keithley_2750
    global myPort
    global Aborted

    try:
        # Channel_selection = Channel_Input.split(',')
        channels = Parse_Channels(Channel_Input)
        devices = channels
        j = 0
        for item in GFETS:
            if j == len(channels):
                break
            elif item.Channel_Num == int(channels[j]):
                devices[j] = item
                j += 1

    except ValueError:
        msg.showinfo()
        msg.showerror("Channel Input", "Malformed Channel Information. Ensure a range of channels is provided. (Ex. 101:105)")
        return()

    if(file_path.cget("text") == "File Path Not Selected"):
        msg.showerror("File Path not selected. Please select, and try again.")
        return()

    start = int(Start_Vg)
    end = int(End_Vg)
    delta = int(Delta_Vg)
    # mode = "a"
    excel_mode = "w"
    Multimeter.Set_Range(Keithley_Range_Entry.get())
    DAC.Update_Gating(start, end, delta)
    DAC.Set_Vg(0)

    if(sweepback_status.get() == 1):
        gate_voltages = DAC.Return_Vg_SweepBack()
    else:
        gate_voltages = DAC.Return_Vg()    
###Matplotlib Section
    LivePlot.clear()
    plot_data = dict()
    # Device_Data = dict()
    data_run = dict()
    for device in devices:
        device_label = device.Device_Name
        plot_data[device_label], = LivePlot.plot([], [], label=device_label)
        # Device_Data[Channel] = ([],[])
        data_run[device_label] = []
    data_run["Gate Voltage"] = []

    
    LivePlot.legend()
    LivePlot.set_xlim(0, End_Vg/1000)
    LivePlot.set_xlabel("Gate Voltage (V)")
    LivePlot.set_ylabel("Resistance (Ohms)")
###End Matplotlib Section

    recorded_data = pd.DataFrame()

    path = file_path.cget("text")
    final_path = path + "/" + file_name.get()

    if(os.path.exists(final_path+".xlsx")):
        excel_mode = "a"
        print('previous file found')
    else:
        excel_mode = "w"


###Physical Chip Measurement Section 
    
    #Stepping through each Gate voltage and measuring resistance of the devices
    for Vg in gate_voltages:
        if(Aborted):
            msg.showinfo("Aborted Measurement!", "Ending on next scan!")
            Aborted = False
            return()
        
        window.title(f"Measuring Dirac Points | Current Gate Voltage:{Vg}")
        DAC.Set_Vg(Vg)
        time.sleep(3)
        resistances = Multimeter.Scan_Channels(",".join(channels))
        #At this point, we have a list of resistance values for the devices, in order that they are supplied
        
        data_run["Gate Voltage"].append(Vg/1000)
        for i, val in enumerate(resistances):
            #Updating the Matplotlib 
            device_label = devices[i].Device_Name
            data_run[device_label].append(val)
            plot_data[device_label].set_xdata(data_run["Gate Voltage"])
            plot_data[device_label].set_ydata(data_run[device_label])

            if(max(resistances)*1.50 > LivePlot.get_ylim()[1]): 
                #Checks if the resistances coming in this time requires a rescaling of the liveplot;
                # it does not rescale if the device is clearly dead (Resistance > 200000 ohms)
                if(max(resistances) > 200000):
                    upper_limit = 100000
                    last = 0
                    for res in resistances:
                        if res < upper_limit and res > last:
                            last = res
    
                    LivePlot.set_ylim(0, last*1.50)
                else:
                    LivePlot.set_ylim(0, max(resistances)*1.50)
        
        fig.canvas.draw()
        fig.canvas.flush_events()

        #End of Awful, Awful code
        resistances.insert(0, Vg/1000)
        print(f"Current Gate Voltage: {Vg} Resistances: {resistances}")
        if(stepup_status.get() == 1):
            msg.showinfo("Done with this step; change resistance")
    
    DAC.Set_Vg(0)

###Physical Chip Measurement Section End


###Packaging into Dataframe Section
    #At this point, we have a dictionary with 1 column of gate voltages, and rest of columns with the resistnace values
    dirac_message = ""
    formatted_data = dict()
    formatted_data["Gate Voltage (V)"] = data_run["Gate Voltage"]
    dirac_dict = dict()
    suffix = channel_suffix_entry.get()
    for keyC in data_run.keys():
        if(keyC != "Gate Voltage"):
            formatted_data[keyC + suffix] = data_run[keyC]
            
            #I'll probably hate myself for this 1 liner, but fuck it, we ball
            dirac_point = data_run["Gate Voltage"][data_run[keyC].index(max(data_run[keyC]))]
            dirac_dict[keyC] = dirac_point
            dirac_message = dirac_message + f"{keyC}:{dirac_point:.2f}|"
    log_textbox.insert(END, dirac_message+"\n")
    #Inserting the gate voltage leftmost column
    recorded_data = pd.DataFrame.from_dict(formatted_data)

###Packaging into Dataframe Section END

###File Export Section
    if(excel_mode == "a"):
        existing_file = pd.read_excel(final_path+".xlsx")
        existing_file.insert(len(existing_file.columns), "","")
        final_output = pd.concat([existing_file, recorded_data], axis=1)
        final_output.to_excel(final_path+'.xlsx', index=False)
    else:
        recorded_data.to_excel(final_path+'.xlsx', index=False)
        
    fig.savefig(final_path + file_name.get() + channel_suffix_entry.get() + ".png")
    msg.showinfo("Measurement Status", "Measurement Finished")
    
    return(dirac_dict)
#File Export Section END

def Check_Device_Status(diracs, connections=False):
    '''
    Looks through diracs provided to determine in a device is shorted, 
        resistance > 200000
    Returns a message that declares if the channel is alive or dead
    ----
    Parameters:
        diracs: a dictionary in the format of {'101':[###], '102':[###], ...}
    '''
    global connected_channels
    connected_channels = list(diracs)
    for i, item in enumerate(connected_channels):
        item = str(item)
        connected_channels[i] = item
    diracs = list(diracs.values())
    connected_msg = ""
    j = 0
    if connections == True:
        for i, device in enumerate(GFETS):
            if i%5 == 0 :
                connected_msg = connected_msg + "\n" + f"WELL {int(1 + i/5)}: "

            if device.Status != 'DEAD':
                if diracs[j] > 200000:
                    GFETS[i].Status = "DEAD"
                else:
                    connected_channels.append(str(device.Channel_Num))
                j += 1
            connected_msg = connected_msg + f" {device.Channel_Num} - {device.Status}  "
        complete_msg ="--------------  Connections Check Complete  --------------\n" + connected_msg
        return(complete_msg)
    else:
        for i, val in enumerate(diracs):
            if val > 200000:
                for j in range(20):
                    if int(connected_channels[i]) == GFETS[j].Channel_Num:
                        GFETS[j].Status = "DEAD"
                        connected_msg = connected_msg + f"   {GFETS[j].Channel_Num}"
                connected_channels[i] = 'DEAD'
        occurances = connected_channels.count('DEAD')
        for i in range(occurances):
            connected_channels.remove('DEAD')
        if connected_msg != "":
            complete_msg = "These devices have died:" + connected_msg 
        else:
            complete_msg = "No devices have died :)"
        return(complete_msg)
   
def Check_Connections_Thread():
    '''
    This thread will run a measurement from 0-0.04V, with steps of 0.02V
    This will print the connected channels and report any devices that have died 
        in the previous run.
    The connected_channels will be updated automatically from the previously selected channels
    '''
    channels = "101:120"
    start_vg = 0
    end_vg = 40
    delta_vg = 20

    t1 = ThreadReturn(target=Begin_Measurement, args=(channels,start_vg,end_vg,delta_vg))
    t1.start()
    diracs = t1.join()
    complete_msg = Check_Device_Status(diracs,True)

    log_textbox.insert(END, complete_msg + "\n")
    log_textbox.insert(END, "Connected channels have been updated. Please proceed to automatic measurements. \n")
    channel_entry.delete(0,END)
    channel_entry.insert(connected_channels)    

def Automatic_Measurement_Thread():
    '''
    Takes the entry from channel_entry input box!!
    Two sweeps @ 0-1.5V, step 0.05V
    2-4 sweeps @ 0-1.5V, step 0.02V 
        This will repeat 4 times, breaking if the difference 
        between diracs between the sweeps is <= 0.04
    After the fourth sweep, a popup will ask if you would like to repeat the 0.02V sweep
    '''
    channels = channel_entry.get()
    start_vg = 0
    end_vg = 1500
    delta_vg = 50

    t1 = ThreadReturn(target=Begin_Measurement,args=(channels,start_vg,end_vg,delta_vg))
    for m in range(2):
        t1.start()
        diracs = t1.join()
        complete_msg = Check_Device_Status(diracs,False)
        log_textbox.insert(END, complete_msg + "\n")
        log_textbox.insert(END)
        channel_entry.delete(0,END)
        channel_entry.insert(connected_channels)
    
    delta_vg = 20
    t1 = ThreadReturn(target=Begin_Measurement,args=(channels,start_vg,end_vg,delta_vg))
    t1.start()
    diracs = t1.join()
    complete_msg = Check_Device_Status(diracs,False)
    log_textbox.insert(END, complete_msg + "\n")
    channel_entry.delete(0,END)
    channel_entry.insert(connected_channels)
    prev_diracs = diracs

    done = False
    for m in range(4):
        if done == True:
            break
        t1.start()
        diracs = t1.join()
        complete_msg = Check_Device_Status(diracs,False)
        log_textbox.insert(END, complete_msg + "\n")
        channel_entry.delete(0,END)
        channel_entry.insert(connected_channels)
        diffs = []
        if len(prev_diracs) == len(diracs):
            for n, (prev_val, new_val) in enumerate(zip(prev_diracs,diracs)):
                diff = new_val - prev_val
                diffs.append(diff)
                if diff > 0.04:
                    done = True
            avg = sum(diffs)/len(diffs)
        prev_diracs = diracs
    repeat = msg.askyesno("Repeat Sweep",f"Would like to repeat the last sweep with Delta = {delta_vg} mV? The average shift between your last two sweeps was {avg}.")
    if repeat == True:
        t1.start()
        diracs = t1.join()
        complete_msg = Check_Device_Status(diracs,False)
        log_textbox.insert(END, complete_msg + "\n")
        channel_entry.delete(0,END)
        channel_entry.insert(connected_channels)
        diffs = []
        if len(prev_diracs) == len(diracs):
            for n, (prev_val, new_val) in enumerate(zip(prev_diracs,diracs)):
                diff = new_val - prev_val
                diffs.append(diff)
                if diff > 0.04:
                    done = True
            avg = sum(diffs)/len(diffs)
        prev_diracs = diracs
        repeat = msg.askyesno(
            "Repeat Sweep",f"Would like to repeat the last sweep with 
            Delta = {delta_vg} mV? The average shift between your last two sweeps was {avg}."
                             )
    else:
        return()

def Abort_Measurement():
    global Aborted
    Aborted = True
#endregion 

#region UI SETUP

window.title("HARDWARE NOT CONNECTED")

channel_label = Label(window, text="Channel Select 101:120")
channel_entry = Entry(window)
channel_suffix_label = Label(window, text="Device Header Suffix")
channel_suffix_entry = Entry(window)

start_vg_label = Label(window, text="Gate Voltage Start (mV)")
start_vg_entry = Entry()
end_vg_label = Label(window, text="Gate Voltage End (mV)")
end_vg_entry = Entry()
delta_vg_label = Label(window, text="Gate Voltage Delta (mV)")
delta_vg_entry = Entry()

sweepback_status = IntVar()
sweepback_checkbox= Checkbutton(variable=sweepback_status)
sweepback_label = Label(window, text="SweepBack?")

stepup_status = IntVar()
stepup_checkbox = Checkbutton(variable=stepup_status)
stepup_label = Label(window, text="Pause Each Delta?")

file_selection_button = Button(window, text="Select File Location", command=Get_File_Path)
file_path = Label(window, text="File Path Not Selected", width=20)
file_name = Entry(window)
file_name.insert(0, "Enter File Name!")

connections_button = Button(window, text="Connections", command=Check_Connections_Thread,)
automatic_button = Button(window, text="Automatic Measurement", command=Automatic_Measurement_Thread)
begin_button = Button(window, text="Begin Measurement", command=Begin_Manual_Measurement_Thread)
abort_button = Button(window, text="Abort Measurement", command=Abort_Measurement)

Keithley_Address_Label = Label(window, text="Keithley Address")
Keithley_Address_Entry = Entry(window)
Keithley_Address_Entry.insert(0,"ASRL4::INSTR")
Keithley_Range_Label = Label(window, text="Measurement Range:")
Keithley_Range_Entry = Entry(window)
Keithley_Range_Entry.insert(0, "10000")
Arduino_Address_Label = Label(window, text="Arduino Address")
Arduino_Address_Entry = Entry(window)
Arduino_Address_Entry.insert(0,"COM3")
Open_Ports_Button = Button(window, text="Open Ports", command=Open_Ports)

Keithley_Address_Label.grid(row=0, column=0)
Keithley_Address_Entry.grid(row=0, column=1)
Arduino_Address_Label.grid(row=0, column=2)
Arduino_Address_Entry.grid(row=0, column=3)
Open_Ports_Button.grid(row=0, column=4)
Keithley_Range_Label.grid(row=0, column=5)
Keithley_Range_Entry.grid(row=0, column=6)

channel_label.grid(row=1,column=0)
channel_entry.grid(row=1, column = 1)
channel_suffix_label.grid(row=1, column=2)
channel_suffix_entry.grid(row=1, column=3)

start_vg_label.grid(row=2, column=0)
start_vg_entry.grid(row=2, column=1)
end_vg_label.grid(row=2, column=2)
end_vg_entry.grid(row=2, column=3)
delta_vg_label.grid(row=2, column=4)
delta_vg_entry.grid(row=2, column=5)

sweepback_label.grid(row=2,column=6)
sweepback_checkbox.grid(row=2, column=7)
stepup_checkbox.grid(row=2, column=9)
stepup_label.grid(row=2, column=8)

file_selection_button.grid(row=3, column=0)
file_path.grid(row=3, column=2)
file_name.grid(row=3, column=1)

begin_button.grid(row=4,column=0)
abort_button.grid(row=4,column=1)
connections_button.grid(row=4, column=2)
automatic_button.grid(row=4, column=3)

log_textbox = Text(master=log_window, width=100)
log_textbox.pack()

canvas = FigureCanvasTkAgg(fig, master=secondary_window)
canvas.draw()
canvas.get_tk_widget().pack()
#endregion

window.mainloop()
