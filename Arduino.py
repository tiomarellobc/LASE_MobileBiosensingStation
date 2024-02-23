import serial

class Arduino:
    def __init__(self, Serial_device, Vg_Start, Vg_end, Vg_delta):
        self.serial_Device = Serial_device
        self.vg_start = Vg_Start
        self.vg_end = Vg_end
        self.Vg_delta = Vg_delta
        self.current_vg = Vg_Start
    
    def Update_Gating(self, start, end, delta):
        """
        Updates the Arduino object's interna
        """
        self.vg_start = start
        self.vg_end = end
        self.Vg_delta = delta
    
    def Return_Gate_Voltages(self):
        Voltages = range(self.vg_start, self.vg_end+self.Vg_delta, self.Vg_delta)
        return(Voltages)
    def Return_Gate_Voltages_SweepBack(self):
        Voltages_Forward = range(self.vg_start, self.vg_end, self.Vg_delta)
        Voltages_Back = range(self.vg_end, self.vg_start, self.Vg_delta)
        return(Voltages_Forward+Voltages_Back)
    def Set_Gate_Voltage(self, Vg_mV):
        self.serial_Device.write(f"V{int(Vg_mV):04d}".encode())