import pyvisa

class DMM:
    
    def __init__(self, VisaInstrument):
        self.insr = VisaInstrument
        self.range = 10000
    
    def Set_Range(self, new_range):
        self.range = new_range
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
    def Write_Message_To_Display(self, Message):
        self.insr.write(f"disp:text:data '{Message}'\n")
        self.insr.write("disp:text:data on\n")
        self.insr.write('trac:cle\n') #Clears the internal buffer
    def Scan_Channels(self, channel_list):
        self.insr.write(f"FUNC 'RES', (@{channel_list})\n") #Sets selected channel range to measure resistance
        #Filter section
        #self.insr.write("res:aver:tcon rep\n")
        #self.insr.write(f"res:aver:tcon coun 10, (@{channel_list})\n")
        #self.insr.write(f"res:aver:stat on, (@{channel_list})\n") #turn off filter
        #There's some weridness here; I suspect that some of our jump issuess may be due to ranging issues
        self.insr.write(f"res:rang {self.range}, (@{channel_list})\n") #Sets the range by suppling an expected resistance value, in this case, 2000 Ohms
        #self.insr.write(f"res:rang:auto on, (@{channel_list})\n") #Sets selected channel range to measure resistance
        #self.insr.write(f"res:rang 1e4, (@{start}:{end})\n") #Sets selected channel range to measure resistance
        #self.insr.write(f"syst:azer:stat off\n")
        #self.insr.write(f"disp:enab off") #Turn off display
        self.insr.write(f"res:nplc 1, (@{channel_list})\n") #Sets integration time to 0.01 => Measure FAST
        self.insr.write(f"rout:scan (@{channel_list})\n") #Sets which chanenls to measure
        self.insr.write('trac:cle\n') #Clears the internal buffer
        self.insr.write('init:cont off\n') #Disables continous intiation
        self.insr.write('trig:sour imm\n') #Sets the trigger for the scan to "immediate"
        self.insr.write(f'trig:coun 1\n') #Measure 1 scan cycle (once over every channel)
        self.insr.write(f'samp:coun {(channel_list.count(","))+1}\n') #Total number of measurements to take (Matches # of channels)
        self.insr.write('rout:scan:tso imm') 
        self.insr.write('rout:scan:lsel int\n') #enables scan
        readings = self.insr.query('read?\n') #Triggers the scans and retrieves the values
        readings = readings.strip() #Removes trailing whitespaces
        readings = readings.replace("\x13", "") #The message from the Keithley 2750 comes with these odd ascii control characters... removed
        readings = readings.replace("\x11", "")
        readings = readings.split(",") #Splits the return message
        readings = list(map(lambda x : float(x), readings)) # Converts the list to a list of floats
        self.Reset()
        return(readings)
