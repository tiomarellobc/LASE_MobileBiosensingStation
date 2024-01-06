import pyvisa
import serial
import time
import csv

class DMM:
    def __init__(self, VisaInstrument):
        self.insr = VisaInstrument
    
    def Get_ID(self):
        Id = self.insr.query("*IDN?\n")
        return(Id)
    def Measure_Resistance(self, channel):
        DMM_Response = self.insr.query(f"measure:resistance? (@{channel})\n")
        Resistance = DMM_Response.split(",")[0]
        return(Resistance)
    def Close_Instrument(self):
        self.insr.close()
    def Reset(self):
        self.insr.write('*RST\n')
        self.insr.write("form:elem read")
        
    def Scan_Channels(self, start, end):
        self.insr.write(f"FUNC 'RES', (@{start}:{end})\n") #Sets selected channel range to measure resistance
        self.insr.write(f"RES:nplc 0.01") #Sets integration time to 0.01 => Measure FAST
        self.insr.write(f"rout:scan (@{start}:{end})\n") #Sets which chanenls to measure
        self.insr.write('trac:cle\n') #Clears the internal buffer
        self.insr.write('init:cont off\n') #Disables continous intiation
        self.insr.write('trig:sour imm\n') #Sets the trigger for the scan to "immediate"
        self.insr.write(f'trig:coun 1\n') #Measure 1 scan cycle (once over every channel)
        self.insr.write(f'samp:coun {(int(end)-int(start))+1}\n') #Total number of measurements to take (Matches # of channels)
        self.insr.write('rout:scan:tso imm') 
        self.insr.write('rout:scan:lsel int\n') #enables scan
        readings = keithley_2750.query('read?\n') #Triggers the scans and retrieves the values
        readings = readings.strip() #Removes trailing whitespaces
        readings = readings.replace("\x13", "") #The message from the Keithley 2750 comes with these odd ascii control characters... removed
        readings = readings.replace("\x11", "")
        readings = readings.split(",") #Splits the return message
        readings = list(map(lambda x : float(x), readings)) # Converts the list to a list of floats
        self.Reset()
        return(readings)
        
class Arduino:
    def __init__(self, Serial_device, Vg_Start, Vg_end, Vg_delta):
        self.serial_Device = Serial_device
        self.vg_start = Vg_Start
        self.vg_end = Vg_end
        self.Vg_delta = Vg_delta
        self.current_vg = Vg_Start
    
    def Update_Gating(self, start, end, delta):
        self.vg_start = start
        self.vg_end = end
        self.Vg_delta = delta
    
    def Return_Gate_Voltages(self):
        Voltages = range(self.vg_start, self.vg_end+self.Vg_delta, self.Vg_delta)
        return(Voltages)
    def Set_Gate_Voltage(self, Vg_mV):
        self.serial_Device.write(f"V{int(Vg_mV):04d}".encode())
   
        


rm = pyvisa.ResourceManager()
print("Devices connected: " + str(rm.list_resources()))

# Open the Keithley 2750
#Setup
keithley_2750 = rm.open_resource('ASRL2::INSTR')
keithley_2750.timeout = 5000
Multimeter = DMM(keithley_2750)
Multimeter.Reset()

print(Multimeter.Get_ID())
print("Multimeter Connected")

myPort = serial.Serial("/dev/cu.usbmodem13201")
time.sleep(2)
print("Arduino Connected")
Gater = Arduino(myPort, 0, 1000, 5)


Running = True
while(Running):
    print("Youre command await")
    UserInput = input()

    if UserInput == "Measure Channel":
        Reading = Multimeter.Measure_Resistance(input("Which Channel"))
        time.sleep(0.1)
        print(Reading)
    elif UserInput == "Set Gating":
        Vg_mV = input("What Gate Voltage (mV)")
        Gater.Set_Gate_Voltage(Vg_mV)
        time.sleep(0.1)

    elif UserInput == "Record Resistances":
        
        start_Channel = int(input("Starting Channel"))
        end_Channel = int(input("Ending Channel"))
        file_name = input("File Name?")
        with open(f"{file_name}.csv", "w", newline='') as file:
            writer = csv.writer(file, dialect="excel")
            Header = []
            Res_Row = []
            for channel in range(start_Channel, end_Channel+1, 1):
                Header.append(channel)
            Resistances = Multimeter.Scan_Channels(start_Channel, end_Channel)
            for val in Resistances:
                Res_Row.append(val)
            writer.writerow(Header)
            writer.writerow(Res_Row)

    
    elif UserInput == "Scan Resistances":
        start = input("Start Channel")
        end = input("EndChannel")
        Resistances = Multimeter.Scan_Channels(start, end)
        time.sleep(0.5)
        print(Resistances)
        

    elif UserInput == "Run Dirac Point Measurement":
        start_vG = int(input("Starting gate Voltage"))
        end_vG = int(input("Ending Gate Voltage"))
        delta_vG = int(input("Delta Gate Voltage"))
        start_Channel = int(input("Beginning Channel"))
        end_Channel = int(input("Ending Channel"))
        file_name = input("File Name?")

        Gater.Update_Gating(start_vG, end_vG, delta_vG)
        Gater.Set_Gate_Voltage(0)
        with open(f"{file_name}.csv", "w", newline="") as file:
            writer = csv.writer(file, dialect='excel')
            Header = ["Gate Voltage"]
            for channel in range (start_Channel, end_Channel+1, 1):
                 Header.append(channel)
            writer.writerow(Header)
            for Vg in Gater.Return_Gate_Voltages():
                Gater.Set_Gate_Voltage(Vg)
                Resistances = Multimeter.Scan_Channels(start_Channel, end_Channel)
                Resistances.insert(0, Vg)
                writer.writerow(Resistances)
                time.sleep(0.01)
            Gater.Set_Gate_Voltage(0)

    
    elif UserInput == "Close All":
        Multimeter.Close_Instrument()
        myPort.close()
        print("Good Bye")
        Running = False
    
    else:
        print("Command not recognized")


