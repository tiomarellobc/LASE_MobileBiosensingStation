class GFET:
    def __init__(self, Channel_Num, Device_Name, Status = "Live", ID = None, Diracs = None):
        self.Channel_Num = Channel_Num
        self.Device_Name = Device_Name
        self.Status = Status
        self.ID = None
        self.Diracs = None
    
    def __str__(self):
        return f"{self.Device_Name} is {self.Status}"
    
    def __repr__(self):
        if self.ID != None:
            return f"({self.Channel_Num}, \'{self.Device_Name}\', \'{self.Status}\', {self.ID}, {self.Diracs})"
        else:
            return f"Channel({self.Channel_Num}, \'{self.Device_Name}\', \'{self.Status}\')"
