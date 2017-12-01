"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     16/08/2011
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2011 

references: sqlite

=====================
Dependencies:
 
wx, 
matplotlib,
sqlite,
scipy,
numpy
=====================
 
==========================
End-User License Agreement:
===========================
This software is created by Intelligent-Infrastructure - Rafal Kucharski (i2) Krakow Polska, who also owns the copyrights. 

By using this software you agree with terms stated below:

1.You can use the software only if You bought it from intelligent-infrastructure, or got written permission of i2 to do so.
2.You can use and modify the software code, as long as you don't sell it's parts commercially.
3.You cannot publish and/or show any parts of the code to third-party users without written permission of i2 
4.If You want to sell the software created by modifying this software, you need to contact with i2 and agree conditions
5.This is one user copy, you cannot use it on multiple computers without written permission to do so
6.You cannot modify this statement
7.You can freely analyze the code, and propose any changes
8. After period defined by special i2 statement this software becomes freeware, so that it can be freely downloaded and/or modified.
9. Parts of this code cannot be used to any other Python software creating without written permission of i2

March 2012, Krakow Poland
"""

import sqlite3, time, sys, os
from random import randint, sample, random, choice
from numpy.random import normal
import matplotlib
matplotlib.interactive(True)
matplotlib.use('WXAgg')
import numpy
import wx.grid
from numpy import percentile, median
from math import sqrt

class DataBase:
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
    ####
    Main class containing DB connection and all methods for DB operations
    """
    def __init__(self, params):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        CONSTRUCTOR:
        
        
        IN: self.Visum - COMobject instance
        IN: self.db_path - wx.filedialog result
        IN: self.initDB - boolean True -> create new database
        IN: self.TSys 
        IN: self.DSeg
        IN: self.Inerpolate
        
        OUT:
        self.con,
        self.cur
        self.con.text_factory = str -string errors
        
        #PROGRESSBAR=TRUE
        """        
        [self.Visum,
         self.db_path,
         self.visum_path,
         self.initDB,
         self.TSys,
         self.DSeg,
         self.Interpolate] = params
         
         
        self.param_dict = {0: [0, "t0", "T0_PRTSYS(" + self.TSys + ")"],
                      1: [1, 'tCur', "TCur_PRTSYS(" + self.TSys + ")"],
                      2: [2, 'Impedance', "Imp_PRTSYS(" + self.TSys + ",AP)"],
                      3: [3, 'Length', "Length"],
                      4: [4, 'AddVal1', "AddVal1"]}
        
        self.dialog = wx.ProgressDialog ('Progress', "Creating new Database", maximum=100)        
        self.con, self.cur = self.Connect_with_DB(self.db_path)
        self.dialog.Update(20)
        self.con.text_factory = str
        if self.initDB == True:
            self.__initialize_DB()
        self.dialog.Destroy()
        
    def __initialize_DB(self):
                """
                ###
                Automatic Plate Number Recognition Support
                (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
                ####
                CONSTRUCTOR CONTINUE, FOR NEW DB:
                
                creates new tables
                imports CL data from VISUM 
                Creates TABLE matrix with initial Values
                
                #PROGRESSBAR=TRUE
                

                #DONE SN: podwojna inicjalizacja - connect? WIESZA SIE JAK INICJALIZUJESZ DO ISTNIEJACEGO PLIKU
                #DONE SN: Overwrite- ptrogram sie wiesza jak baza danych juz istnieje - rozwiazac albo poprzez drop tables,
                 albo proprzez blokade mozliwosci nadpisania, 
                 albo poprzez usuniecie pliku.
                """
                try:
                    self.dialog.Update(30)
                    self.__create_DB_Tables()
                    self.dialog.Update(40)
                    self.__insert_CLs_to_DB()
                    self.dialog.Update(60)
                    self.__insert_default_Matrix_to_DB()
                    self.dialog.Update(80)
                except:
                    caption = "Database already exists, initialize corrupted"
                    dlg = wx.MessageDialog(self, caption, "i2 APNR", wx.OK)
                    dlg.Destroy()
    
    def __init_Visum(self, path=None):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        NOT USED HERE        
        """
        import win32com.client        
        self.Visum = win32com.client.Dispatch('Visum.Visum')
        if path != None: self.Visum.LoadVersion(path)
        return self.Visum
    
    def Connect_with_DB(self, path):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Connect with database
        """
        
        con = sqlite3.connect(path)
        cur = con.cursor()
        con.text_factory = str
            #TD, wrzucic statystyki bazy danych na pierwszy panel: liczba tabel, liczba rekordow w tabelach, nazwa sciezki, data stworzenia,etc"
        return con, cur 
    
    def __create_DB_Tables(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Create DataBase Tables
        """
        
        self.cur.execute("""create table CountLocations(No INT PRIMARY KEY,
                                        CLCode VARCHAR, 
                                        WKTLoc VARCHAR,
                                        FromNodeNo INT,
                                        ToNodeNo INT,
                                        RelPos DECIMAL(2,0),
                                        Link_Length DECIMAL(2,0), 
                                        ReverseCode VARCHAR,
                                        VOL INT,
                                        VOL_ERROR INT,
                                        VOL_FRATAR_FROM DECIMAL(2,0),
                                        VOL_FRATAR_TO DECIMAL(2,0))""")
        self.cur.execute("""create table Matrix(IdD INTEGER PRIMARY KEY,
                                        FromCLCode VARCHAR, 
                                        ToCLCode VARCHAR,
                                        State VARCHAR,
                                        VOLUME_VISUM INT, 
                                        T0 DECIMAL(2,0),
                                        TCur DECIMAL(2,0),
                                        Imp DECIMAL(2,0),
                                        DIST DECIMAL(2,0),
                                        APNR_VOLUME_OD INT, 
                                        APNR_VOLUME_DETECTED INT,
                                        APNR_VOLUME_ANY INT,
                                        APNR_VOLUME_ERROR INT,
                                        APNR_TMIN_OD DECIMAL(2,0),
                                        APNR_TMEAN_OD DECIMAL(2,0),
                                        APNR_TMOD_OD DECIMAL(2,0),
                                        APNR_TMAX_OD DECIMAL(2,0),
                                        APNR_TMIN_DETECTED DECIMAL(2,0),
                                        APNR_TMEAN_DETECTED DECIMAL(2,0),
                                        APNR_TMOD_DETECTED DECIMAL(2,0),
                                        APNR_TMAX_DETECTED DECIMAL(2,0),
                                        APNR_TMIN_ANY DECIMAL(2,0),
                                        APNR_TMEAN_ANY DECIMAL(2,0),
                                        APNR_TMOD_ANY DECIMAL(2,0),
                                        APNR_TMAX_ANY DECIMAL(2,0),
                                        APNR_VOLUME_FRATAR INT,
                                        PATHNODES VARCHAR,
                                        CONTAINS_IDD VARCHAR,
                                        IS_CONTAINED_IN_IDD VARCHAR,
                                        enabled VARCHAR,
                                        mint INT,
                                        maxt INT)""")
        
        self.cur.execute("""create table DetectedVehicles(IdD INTEGER PRIMARY KEY, 
                                        CLCode INT, 
                                        DetectionTime INT,
                                        DetectionTimeIP INT, 
                                        VehType VARCHAR, 
                                        PlateNo)""")
        self.con.commit()
    
    def __insert_CLs_to_DB(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Pobierz CountLocations z parametrami z Visuma i wstaw do Bazy danych
        """
        try:            
            self.Visum.Net.CountLocations.AddUserDefinedAttribute("i2_APNR_VOL_FRATAR_TO","i2_APNR_VOL_FRATAR_TO","i2_APNR_VOL_FRATAR_TO",2)
            self.Visum.Net.CountLocations.AddUserDefinedAttribute("i2_APNR_VOL_FRATAR_FROM","i2_APNR_VOL_FRATAR_FROM","i2_APNR_VOL_FRATAR_FROM",2)
        except:
            pass
        CL = self.Visum.Net.CountLocations.GetMultipleAttributes(["No", "Name", "WKTLoc", "FromNodeNo", "ToNodeNo", "RelPos", "Link\Length", "LINK\REVERSELINK\CONCATENATE:COUNTLOCATIONS\Name","i2_APNR_VOL_FRATAR_FROM","i2_APNR_VOL_FRATAR_TO"])
        self.cur.executemany("insert into CountLocations(No,CLCode,WKTLoc,FromNodeNo,ToNodeNo,RelPos,Link_Length,ReverseCode,VOL_FRATAR_FROM,VOL_FRATAR_TO )values(?,?,?,?,?,?,?,?,?,?)", CL)
        self.con.commit()
    
    def set_Paths(self,Paths):
        self.Paths=Paths
        
    def __updateConsole(self, flag, t=None): 
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        export tylko do pliku
        
        """
        PK=True
        if not PK:
            
            flag += "\n"
            
            
            filereport= open(self.Paths["Report"], 'a')
            filereport.write(flag)
            filereport.close()
            
    def Get_Path_Cost(self, FromCL, ToCL, typ):
            """
            ###
            Automatic Plate Number Recognition Support
            (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
            ####
            Gets shortest path cost from Visum. Path between FromCL and ToCL
            
            IN: FromCL, ToCL, typ, self.param_dict
            
            self.param_dict = {0: [0,"t0","T0_PRTSYS("+self.TSys+")"],
                      1: [1,'tCur',"TCur_PRTSYS("+self.TSys+")"],
                      2: [2,'Impedance',"Imp_PRTSYS("+self.TSys+",AP)"],
                      3: [3,'Length',"Length"],
                      4: [4,'AddVal1',"AddVal1"]}
            
            OUT: [COST,STATE]
            
            if typ==-1:
                returns [0,NodeChain] (NodeNos)
            
            STATE: 
            a) ok - path exists + shortest path between 2CLs crosses them naturally
            b) diag - FromCl = ToCL
            c) reverse - FromCL LinkNo = ToCL LinkNo - countlocations on opposite sides of the road
            d) no SP found 
            e) loops, see below:
                
              NodeA        CL1        Node B                        Node C           CL3          Node D
                |          \/           |                             |               \/            |
                X<======================X<============================X<============================X
                X======================>X============================>X============================>X
                           /\                                                        /\
                          CL2                                                        CL4
            \ ToCL | 1        | 2       | 3          |4
            FromCL |__________|_________|____________|___________
            1      |diag      |reverse  | both_loop  | tail_loop
            2      |reverse   |diag     | head_loop  | ok
            3      |ok        |head_loop| diag       | reverse
            4      |tail_loop |both_loop| reverse    | diag
            """
            
            def Obetnij_Koncowki(Cost):
                """
                ###
                Automatic Plate Number Recognition Support
                (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
                ####
                If state == ok, obetnij koncowki.
                """
                
                poczatek = self.Visum.Net.Links.ItemByKey(FromCL[3], FromCL[4]).AttValue(Link_Attr) * float(FromCL[5])
                koniec = self.Visum.Net.Links.ItemByKey(ToCL[3], ToCL[4]).AttValue(Link_Attr) * (1 - float(ToCL[5]))
                try:
                    Cost = Cost - poczatek - koniec
                except:
                    Cost = int(Cost[:-1]) - poczatek - koniec #10s -> 10
                return Cost
            
            
            self.param_dict = {0: [0, "t0", "T0_PRTSYS(" + self.TSys + ")"],
                      1: [1, 'tCur', "TCur_PRTSYS(" + self.TSys + ")"],
                      2: [2, 'Impedance', "Imp_PRTSYS(" + self.TSys + ",AP)"],
                      3: [3, 'Length', "Length"],
                      4: [4, 'AddVal1', "AddVal1"]}
            
                  
            if typ == -1:
                typ = 0
                RSearch = self.Visum.Analysis.RouteSearchPrT 
                [SP_crit, SP_attr, Link_Attr] = self.param_dict[typ]
                List = self.Visum.Lists.CreatePrTPathSearchLegList 
                List.AddKeyColumns() 
                List.AddColumn(SP_attr, 3, 2)
                Container = self.Visum.CreateNetElements()
                Container.Add(self.Visum.Net.Nodes.ItemByKey(FromCL[3]))
                Container.Add(self.Visum.Net.Nodes.ItemByKey(ToCL[4]))
                RSearch.Clear()
                
                RSearch.Execute(Container, self.TSys, SP_crit)
                del Container
                NodeChain = RSearch.NodeChainPrT
                Nodes = []
                for i in range(NodeChain.Count):
                    Nodes.append(NodeChain.Item(i + 1).AttValue("No"))
                return [0, Nodes]
                    
                
  
                
                
            RSearch = self.Visum.Analysis.RouteSearchPrT #init route search object
            [SP_crit, SP_attr, Link_Attr] = self.param_dict[typ]
            List = self.Visum.Lists.CreatePrTPathSearchLegList #init list
            List.AddKeyColumns() #nie wiem, ale podobno ok
            List.AddColumn(SP_attr, 3, 2) #dodaj interesujacy parametr
            
            Cost = -1
            state = 'ok'  
            #exception1: CLs on the same road  
            if FromCL[1] == ToCL[7]:
                state = "reverse" 
            #exception2: From = To
            elif FromCL[1] == ToCL[1]:
                state = "diag" 
            else:
                Container = self.Visum.CreateNetElements()
                Container.Add(self.Visum.Net.Nodes.ItemByKey(FromCL[3]))
                Container.Add(self.Visum.Net.Nodes.ItemByKey(ToCL[4]))
                RSearch.Clear()
                RSearch.Execute(Container, self.TSys, SP_crit)
                del Container
                NodeChain = RSearch.NodeChainPrT
                #exception6: null SP
                if NodeChain.Count == 0:                    
                    state = "null path" 
                #exceptions: shortcuts 
                else:
                    #exception3: shortcut at begining
                    if NodeChain.Item(2).AttValue("No") != FromCL[4]:
                        state = "head loop" 
                    if NodeChain.Item(NodeChain.Count - 1).AttValue("No") != ToCL[3]:
                        #exception4: shortcut at begining and at the end
                        if state == "head loop":
                            state = "both loop" 
                        #exception5: shortcut at the end
                        else:
                            state = "tail loop" 
                #exception6: no SP found
                if state != "null path":
                    try:
                        Cost = List.SaveToArray(1, 1)[0][1]
                    except:
                        Cost = -1
                        state = "no SP found" 
            if state == "ok":
                return [Obetnij_Koncowki(Cost), state]
            else:
                if SP_crit < 2: #14s ->14
                    try:
                        Cost = int(Cost[:-1])
                    except:
                        pass
                return [Cost, state]
    
    def Populate_Matrix_from_Visum(self, typ, i=0):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Execute smikmatrix between CLs.
        
        IN: typ
        IN: i
        
        typ             :-1            0    1     2    3    4
        col              :PATHNODES     T0   TCur  IMP  DIST Volume_Visum

        if typ==4: self.Get_Visum_Volume
        else: self.Get_Path_Cost
        
        if i=0: Get_Path_Cost returns Skim Mtx value
        if i=1: Get_Path_Cost returns "STATE"
        see Get_Path_Cost
        
        Query="UPDATE Matrix SET COL = ? where FromCLCode= FromCL[1] and ToCLCode= ToCL[1]"
                
        
        """
        cols = ["T0", "TCur", "Imp", "DIST", "Volume_Visum"]
        self.cur.execute("select * from CountLocations")
        CLs = self.cur.fetchall()
        if i == 0:
            col = cols[typ]
        elif typ == -1:
            col = "PATHNODES"            
        else:            
            col = "STATE"
        noCLs = len(CLs)
        self.dialog = wx.ProgressDialog ('Progress', "Visum skim matrix Calculations", maximum=noCLs + 1)
        j = 0
        RSearch = self.Visum.Analysis.RouteSearchPrT
        try:           
            RSearch = self.Visum.Analysis.RouteSearchPrT
            Container = self.Visum.CreateNetElements()
            Container.Add(self.Visum.Net.Nodes.ItemByKey(CLs[0][3]))
            Container.Add(self.Visum.Net.Nodes.ItemByKey(CLs[0][4]))
            RSearch.Execute(Container, self.TSys, 0)
        except:
            self.dialog.Destroy()
            wx.MessageBox("Please check TSys", "i2 APNR Error", style=wx.OK | wx.ICON_ERROR)
            return
        for FromCL in CLs:
            j += 1            
            self.dialog.Update(j)
            for ToCL in CLs:
                if col == "Volume_Visum":
                    row = self.Get_Visum_Volume([FromCL, ToCL])
                else:
                    row = self.Get_Path_Cost(FromCL, ToCL, typ)[i]
                
                Query = "UPDATE Matrix SET " + col + " = '" + str(row) + "' where FromCLCode= '" + FromCL[1] + "' and ToCLCode= '" + ToCL[1] + "'"
                self.con.execute(Query)
        self.dialog.Destroy()
        self.con.commit()

    
    def XLS_to_DB(self, path):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        #DONE SN: dokumentacja!
        
        in: path of folder with txt files
        out: updated database
        
        Procedure lists all files with extension txt in selected path
        Afterwards, it process each file by reading each line to the end of file
        First of all, it distinguish type of information in each line by first character:
        - if it's '*' - this line is considered as timestamp
        the following part is considered as begining of time interval
        - it it's 'l' or 'c' this line is considered as information of detected vehicle (light weight truck or car),
        the following part is considered as plate number of detected vehicle
        
        After each two timestamps interpolated detection time is calculated for each vehicle in the analysing time period
        

        #PROGRESSBAR=TRUE 
        
        """
        vehtype_slownik={
                         "":"SO",
                         " ":"SO",
                         " SD":"SD",
                         "A":"A",
                         "AUT":"A",
                         "AUTIBUS":"A",
                         "AUTOBUS":"A",
                         "AUTOKAR":"A",
                         "B":"BUS",
                         "BUS":"BUS",
                         "BUC":"BUS",
                         "BUD":"BUS",
                         "CP":"SCP",
                         "CS":"SC",
                         "DS.":"SD",
                         "O":"SO",
                         "S.C.":"SC",
                         "SC":"SC",
                         "SCN":"SCP",
                         "SCP":"SCP",
                         "SD":"SD",
                         "SD ZAGRANICZNY":"SD",
                         "SDP":"SD",
                         "SO":"SO",
                         "SP":"SO",
                         "ZAGRANICZNY":"SO"        
                         }
                         
                        
        
        import xlrd
        ListOfFilePaths = []
        ListOfFilenames = []        
        for dirname, dirnames, filenames in os.walk(path):
            for filename in filenames:
                if filename[-4:] in ['.xls','xlsx']:
                    if filename[0] in ["B","O","G","L"]:
                        ListOfFilePaths.append(os.path.join(dirname, filename))
                        ListOfFilenames.append(filename)                    
               
        self.dialog = wx.ProgressDialog ('Progress', "Importing data from txt files", maximum=len(ListOfFilenames))
        self.con.commit()        
        u = 0
        for ind,filename in enumerate(ListOfFilePaths):            
            
            plik = xlrd.open_workbook(filename)                       
            arkusz = plik.sheet_by_name('Arkusz1')            
            tablica="cos" 
            Wyroznik=ListOfFilenames[ind][:2]         
            CLCode=ListOfFilenames[ind][2:5]
            print filename
            print CLCode            
            for rownum in range(10,arkusz.nrows): 
                vehtype=arkusz.cell(rownum,3).value
                vehtype.replace(" ", "")                                                                   
                godzina=arkusz.cell(rownum,1).value
                try:                                       
                    tablica=Wyroznik+str(arkusz.cell(rownum,2).value)
                except:
                    message=filename+"; numer wiersza "+str(rownum)+"; tablica"+str(tablica)+"; godzina "+str(godzina)
                    wx.MessageBox(message, "i2 APNR Error", style=wx.OK | wx.ICON_ERROR)                                
                if vehtype=="" and godzina=="" and tablica =="":
                    break               
                try:
                    tablica=tablica.upper()
                except:
                    pass
                try:                    
                    vehtype=vehtype.upper()                    
                except:
                    message=filename+"; numer wiersza "+str(rownum)+"; tablica"+str(tablica)+"; godzina "+str(godzina)
                    wx.MessageBox(message, "i2 APNR Error", style=wx.OK | wx.ICON_ERROR)
                try:                
                    tupla=xlrd.xldate_as_tuple(godzina,0)                
                    godzina=tupla[3]*3600+tupla[4]*60 
                except:
                    message=filename+"; numer wiersza "+str(rownum)+"; tablica"+str(tablica)+"; godzina "+str(godzina)
                    wx.MessageBox(message, "i2 APNR Error", style=wx.OK | wx.ICON_ERROR)                    
                try:
                    vehtype=vehtype_slownik[vehtype]
                except:
                    message="Nowy typ pojazdu "+str(vehtype)
                    wx.MessageBox(message, "i2 APNR Error", style=wx.OK | wx.ICON_ERROR)                    
                
                               
                wrzutka=[CLCode,godzina,godzina,vehtype,tablica]
                self.cur.execute("""insert into DetectedVehicles(
                                    ClCode,DetectionTime,DetectionTimeIP,VehType,PlateNo) values (?,?,?,?,?)""", wrzutka)
                
                
            self.con.commit()
            u += 1
            self.dialog.Update(u)            
        self.dialog.Destroy()
        
                
 
    
    def Txt_to_DB2(self, path):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        #DONE SN: dokumentacja!
        
        in: path of folder with txt files
        out: updated database
        
        Procedure lists all files with extension txt in selected path
        Afterwards, it process each file by reading each line to the end of file
        First of all, it distinguish type of information in each line by first character:
        - if it's '*' - this line is considered as timestamp
        the following part is considered as begining of time interval
        - it it's 'l' or 'c' this line is considered as information of detected vehicle (light weight truck or car),
        the following part is considered as plate number of detected vehicle
        
        After each two timestamps interpolated detection time is calculated for each vehicle in the analysing time period
        

        #PROGRESSBAR=TRUE 
        
        """
        ListOfFilePaths = []
        ListOfFilenames = []
        Error = False
        for dirname, dirnames, filenames in os.walk(path):
            for filename in filenames:
                if filename[-4:] == '.txt':
                    ListOfFilePaths.append(os.path.join(dirname, filename))
                    ListOfFilenames.append(filename)
                    
                else:
                    Error = True
        self.dialog = wx.ProgressDialog ('Progress', "Importing data from txt files", maximum=len(ListOfFilenames))
        '''           
        if Error:
            self.ErrMsg('there are non-txt files in directory with results')
        '''
        self.con.commit()
        
        u = 0
        for ind, filename in enumerate(ListOfFilePaths):
            file = open(filename)
            Interval = [None, None]
            DBIndex = 0
            Results = []
            LineInd = 0
            while 1:
                line = file.readline()
                LineInd = LineInd + 1
                if not line:
                    ResultsT = [tuple(r) for r in Results]                    
                    self.cur.executemany('insert into DetectedVehicles(ClCode,DetectionTime,DetectionTimeIP,VehType,PlateNo) values (?,?,?,?,?)', ResultsT)
                    self.con.commit()
                    
                    break

                if line[0] == '*':

                    h = int(line[1:3])
                    m = int(line[4:6])     
                    try:
                        sec=int(line[7:9])
                    except:
                        sec=0                              
                    time = 60 * (60 * h + m) + sec #INCLUDE SECONDS
                    Interval[0] = [LineInd, time]
                    if Interval[1] != None:
                        TimeInt = Interval[0][1] - Interval[1][1]
                        IndexChange = Interval[0][0] - Interval[1][0]
                        if IndexChange <= 1:
                            pass
                        else:
                            TimeDelta = TimeInt / (IndexChange - 1)
                            TimeCh = []
                            for i in range(1, IndexChange):
                                Results[DBIndex - i][2] = Interval[0][1] - i * TimeDelta                                
                            '''
                            zapisz czasy w bazie danych
                            '''    
                    Interval[1] = Interval[0]
                    Interval[0] = None
                    
                else:
                    if line[0:2] == 'l-' or line[0:2] == 'L-' :
                        typ = 'LKW'
                        plateno = line[2:-1]
                        plateno = plateno.upper()
                    elif line[0:2] == ' p' or line[0:2] == ' P' :
                        typ = 'Car'                        
                        plateno = line[2:-1]
                        plateno = plateno.upper()
                    else:
                        typ = 'Car'                        
                        plateno = line[0:-1]
                        plateno = plateno.upper()
                        
                    try:
                        time = Interval[1][1]
                    except:
                        print Interval
                        print line
                        print filename
                        

                    CLCode = ListOfFilenames[ind][:-4]
                    Results.append([str(CLCode), time, time, typ, str(plateno)])
                    DBIndex = DBIndex + 1                
            u += 1
            self.dialog.Update(u)            
        self.dialog.Destroy()
    
                
    def __insert_default_Matrix_to_DB(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        
        Procedure to Get Visum SKIM Mtx from set of CountLocations
        OUT:
        table Matrix
        
        """
        self.cur.execute("select * from CountLocations")
        CLs = self.cur.fetchall()
        for FromCL in CLs:
            for ToCL in CLs:
                row = [FromCL[1], ToCL[1],'yes',0,99999999]
                self.con.execute("insert into Matrix(FromCLCode,ToCLCode,enabled,mint,maxt)  values(?,?,?,?,?)", row)
        self.con.commit()
           
    def Make_Paths(self, j=0):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        CREATES VISUM PATHS BETWEEN CL pairs
        
        Creates path only if SP found & SP NodeChain len > 0
        
        destroys pathsets 12 & 1212 in Visum
        
        IN: j - criteria from self.param_dict
        
        self.param_dict = {0: [0,"t0","T0_PRTSYS("+self.TSys+")"],
                      1: [1,'tCur',"TCur_PRTSYS("+self.TSys+")"],
                      2: [2,'Impedance',"Imp_PRTSYS("+self.TSys+",AP)"],
                      3: [3,'Length',"Length"],
                      4: [4,'AddVal1',"AddVal1"]}
                      
        OUT:Pathset = 12 -> Regular Paths
        OUT:Pahtset = 1212 -> reverse (O/D trips)
        
        
        #TO DO RK: W wolnej chwili,pobrac od razu z bazy 3 kolumny i zrobic update po calosci, bedzie szybciej.
        """
       
        def Add_Path(FromCL, ToCL, Visum, i, Vol, PathSetNo):            
            RSearch = self.Visum.Analysis.RouteSearchPrT 
            RSearch.Clear()
            [SP_crit, SP_attr, Link_Attr] = self.param_dict[j]
            self.__updateConsole("Parametry Add_Path: SP_crit %s SP_Attr %s LinkAttr %s "%(SP_crit, SP_attr, Link_Attr))
                    
            Container = self.Visum.CreateNetElements()
            Container.Add(self.Visum.Net.Nodes.ItemByKey(FromCL[3]))
            Container.Add(self.Visum.Net.Nodes.ItemByKey(ToCL[4]))
            self.__updateConsole("Dodano dwa wezly FromCL %s ToCL %s "%(FromCL[3], ToCL[4]))
            try:
                RSearch.Execute(Container, self.TSys, SP_crit)
                self.__updateConsole("Udalo sie znalezc sciezke: Tsys %s SP_crit %s "%(self.TSys, SP_crit))                
            except:  
                self.__updateConsole("Nie udalo sie znalezc sciezki: Tsys %s SP_crit %s "%(self.TSys, SP_crit))                        
                return i
            
            NodeChain = RSearch.NodeChainPrT
            if NodeChain.Count > 0:
                i += 1
                try:
                    pathno=int(FromCL[1])*100000+int(ToCL[1])
                except:
                    pathno=i
                
                Visum.Net.AddPath(pathno, 12, 0, 0, NodeChain)                
                Visum.Net.Paths.ItemByKey(PathSetNo, pathno).SetAttValue("Vol", Vol)
                self.__updateConsole("Udalo dodac sciezke: pathno %s Vol %s "%(pathno,Vol))
                RSearch.Clear()                
                return i
            else:
                self.__updateConsole("Nie udalo sie dodac - tylko jeden wezel")
                return i
            
        self.param_dict = {0: [0, "t0", "T0_PRTSYS(" + self.TSys + ")"],
                      1: [1, 'tCur', "TCur_PRTSYS(" + self.TSys + ")"],
                      2: [2, 'Impedance', "Imp_PRTSYS(" + self.TSys + ",AP)"],
                      3: [3, 'Length', "Length"],
                      4: [4, 'AddVal1', "AddVal1"]}  
        self.__updateConsole("poczatek DB.Make_Paths")   
        try:
            
            self.Visum.Net.RemovePathSet(self.Visum.Net.PathSets.ItemByKey(12))
            self.Visum.Net.RemovePathSet(self.Visum.Net.PathSets.ItemByKey(1212))
            self.__updateConsole("udalo sie usunac stare sciezki")
        except:
            self.__updateConsole("nie udalo sie usunac starych sciezek")
            
           
        self.Visum.Net.AddPathSet(12)
        self.Visum.Net.AddPathSet(1212)
        self.__updateConsole("stworzono nowe zbiory sciezek")
        self.cur.execute("select * from CountLocations")
        CLs = self.cur.fetchall()
        self.__updateConsole("pobrano %s CL"%(len(CLs)))
        self.dialog = wx.ProgressDialog ('Progress', "Creating PrT paths in Visum", maximum=len(CLs) + 1)
        flag = 0
        i = 0
        u = 0
        for FromCL in CLs:
            u += 1
            for ToCL in CLs:
                self.dialog.Update(u)
                if FromCL[1] != ToCL[1]:                    
                    Query = 'select APNR_VOLUME_OD from Matrix where FromCLCode= ? and ToCLCode= ?'
                    FilterResult = self.cur.execute(Query, (FromCL[1], ToCL[1])).fetchall()[0][0]
                    self.__updateConsole("FilterResult="+str(FilterResult))
                    try: 
                        int(FilterResult)
                    except:
                        FilterResult=0
                    if FromCL[1] == ToCL[9]:                        
                        i = Add_Path(FromCL, ToCL, self.Visum, i, FilterResult, 1212)
                    self.__updateConsole("Wchodze do AddPath z parametrami: FromCL %s ToCL %s i %s FilterResult %s"%(FromCL, ToCL, i, FilterResult))
                    i = Add_Path(FromCL, ToCL, self.Visum, i, FilterResult, 12)
                    self.__updateConsole("Wychodze z AddPath z parametrami: FromCL %s ToCL %s i %s FilterResult %s"%(FromCL, ToCL, i, FilterResult))
                    
                       
                    
        self.dialog.Destroy()
    
    def insert_CLStatistics_to_DB(self):
            """
            
            ###
            Automatic Plate Number Recognition Support
            (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
            ####
            NIE UZYWANE ???
            
            """
            '''
            zapelnia baze danych clstatistics na podstawie danych z detectedvehicles i countlocations
            zczytaj wszystkie kody CL
            '''
            CLCodesI = self.cur.execute('select CLCode from CountLocations').fetchall()
            CLCodes = []
            for i in range(len(CLCodesI)):
                CLCodes.append(str(CLCodesI[i][0]))
            TableOfComb = []
            '''
            wszystkie pojedyncze CL, kombinacje par i trojek
            '''
            for Code1 in CLCodes:
                TableOfComb.append([Code1])                
            for Code1 in CLCodes:
                for Code2 in CLCodes:
                    if Code1 == Code2:
                        pass
                    else:
                        TableOfComb.append([Code1, Code2])
            for Code1 in CLCodes:
                for Code2 in CLCodes:
                    if Code1 == Code2:
                        pass
                    else:
                        for Code3 in CLCodes:
                                if Code3 == Code1 or Code3 == Code2:
                                    pass
                                else:
                                    TableOfComb.append([Code1, Code2, Code3])
            '''
            dla kazdego elementu zastosuj filtr ze statystykami, zapisz do bazy danych
            '''


            
                
            for T in TableOfComb:
                Vol, DT, TT = self.Filter(True, T)
                TT = str(T)
                TT.replace('[', '')
                TT.replace(']', '')
                self.cur.execute('insert into CLStatistics(ClCodes,CLType,Volume,DetectionTimes,TravelTimes) values (?,?,?,?,?) ', (TT, len(T), Vol, str(DT), str(TT)))
            self.con.commit()
            
    def Filter(self, stats=False, CountLocationNo=None, VehType=None, PlateNo=None, FromTime=0, ToTime=1000000):
        """
        
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        MAIN DB Query SUPPORT:
        
        Creates DetectedVehiclesTemp table
        
        Procedure filters Detected Vehicles database depending on selected parameters in GUI
        It filters by sequence of count locations (up to 3), vehicle type, specified plate number and timerange
        
        The output is usually: Plate Number,Vehicle Type,DetectionTime on CL1,DetectionTime on CL2, DetectionTime on CL3
        but if there are no count location sequence in input data the output is: Plate Number,Vehicle Type,Count Location Code
        
        Because of problems of time consuming inner join in sqlite3 database over 100 000 records, the database Detected Vehicles is narrowed down by sql basic operations
        

        #DONE SN: opisz zasady wyrzucania filtra, 
        #DONE SN: filtr po tablicy - inny format wyrzucenia danych, niech poda tez CL#
        
        """
        
        if CountLocationNo == None:
            CountLocationNo = []
        try:
            self.cur.execute('drop table DetectedVehiclesTemp')
        except:
            pass
        try:            
            self.cur.execute("""create table DetectedVehiclesTemp(IdD INTEGER PRIMARY KEY, 
                                        CLCode INT, 
                                        DetectionTime INT,
                                        DetectionTimeIP INT, 
                                        VehType VARCHAR, 
                                        PlateNo VARCHAR)""")

        except:
            pass

        Results = []
        if CountLocationNo == None or len(CountLocationNo) == 0:
            operator_CL = ' <> '
            var_CL = '0'
        elif len(CountLocationNo) == 1:
            operator_CL = ' = '
            var_CL = str(CountLocationNo[0])
        else:
            var_CL = str(tuple(CountLocationNo))
            operator_CL = ' in '
                    
        if VehType == None:
            operator_VT = ' <> '
            var_VT = ''
        else:
            operator_VT = ' = '
            var_VT = VehType
        
        if PlateNo == None:
            operator_PL = ' <> '
            var_PL = ''
        else:
            if '%' in PlateNo:
                operator_PL = ' like '
                var_PL = PlateNo 
            else:
                operator_PL = ' = '
                var_PL = PlateNo
        #TO DO SN: Nie wyrzuca dobrze CLcode dla tablicy rejestracyjnej. (linijka else)
        if len(CountLocationNo) == 1 or len(CountLocationNo) == 0:
            if self.Interpolate:
                if PlateNo == None:
                    Results = self.cur.execute('select PlateNo,VehType,DetectionTimeIP from DetectedVehicles where CLCode' + operator_CL + '? and PlateNo' + operator_PL + '? and (DetectionTimeIP between ? and ?) and VehType' + operator_VT + '?' + ' order by DetectionTime', (var_CL, str(var_PL), FromTime, ToTime, var_VT)).fetchall()
                else:
                    Results = self.cur.execute('select PlateNo,VehType,DetectionTimeIP,CLCode from DetectedVehicles where CLCode' + operator_CL + '? and PlateNo' + operator_PL + '? and (DetectionTimeIP between ? and ?) and VehType' + operator_VT + '?' + ' order by DetectionTime', (var_CL, str(var_PL), FromTime, ToTime, var_VT)).fetchall()

            else:
                if PlateNo == None:
                    Results = self.cur.execute('select PlateNo,VehType,DetectionTime from DetectedVehicles where CLCode' + operator_CL + '? and PlateNo' + operator_PL + '? and (DetectionTime between ? and ?) and VehType' + operator_VT + '?' + ' order by DetectionTime', (var_CL, str(var_PL), FromTime, ToTime, var_VT)).fetchall()
                else:
                    Results = self.cur.execute('select PlateNo,VehType,DetectionTime,CLCode from DetectedVehicles where CLCode' + operator_CL + '? and PlateNo' + operator_PL + '? and (DetectionTime between ? and ?) and VehType' + operator_VT + '?' + ' order by DetectionTime', (var_CL, str(var_PL), FromTime, ToTime, var_VT)).fetchall()
                
            
            
            
        if len(CountLocationNo) == 2:
            self.cur.execute('Insert into DetectedVehiclesTemp(CLCode,DetectionTime,DetectionTimeIP,VehType,PlateNo) select CLCode,DetectionTime,DetectionTimeIP,VehType,PlateNo from DetectedVehicles where CLCode' + operator_CL + var_CL + ' and PlateNo' + operator_PL + '? and (DetectionTime between ? and ?) and VehType' + operator_VT + '?' + 'and PlateNo <> "-" ', (str(var_PL), FromTime, ToTime, var_VT))
            self.con.commit()
            if self.Interpolate:
                Results = self.cur.execute('''
                    select dvPoint1.PlateNo, dvPoint1.VehType, dvPoint1.DetectionTimeIP, dvPoint2.DetectionTimeIP
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTimeIP < dvPoint2.DetectionTimeIP 
                    order by dvPoint1.DetectionTimeIP''', (str(CountLocationNo[0]), str(CountLocationNo[1]))).fetchall()
            else:
                Results = self.cur.execute('''
                    select dvPoint1.PlateNo, dvPoint1.VehType, dvPoint1.DetectionTime, dvPoint2.DetectionTime
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTime < dvPoint2.DetectionTime 
                    order by dvPoint1.DetectionTime''', (str(CountLocationNo[0]), str(CountLocationNo[1]))).fetchall()
        if len(CountLocationNo) == 3:
            self.cur.execute('Insert into DetectedVehiclesTemp(CLCode,DetectionTime,DetectionTimeIP,VehType,PlateNo) select CLCode,DetectionTime,DetectionTimeIP,VehType,PlateNo from DetectedVehicles where CLCode' + operator_CL + var_CL + ' and PlateNo' + operator_PL + '? and (DetectionTime between ? and ?) and VehType' + operator_VT + '?' + 'and PlateNo <> "-" ', (str(var_PL), FromTime, ToTime, var_VT))
            self.con.commit()
            if self.Interpolate:
                ResultsT = self.cur.execute('''
                    select dvPoint1.PlateNo
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo 
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTime < dvPoint2.DetectionTime
                    intersect
                    select dvPoint1.PlateNo
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo 
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTimeIP < dvPoint2.DetectionTimeIP
                    ''', (str(CountLocationNo[0]), str(CountLocationNo[1]), str(CountLocationNo[1]), str(CountLocationNo[2]))).fetchall()
            
                PlateList = tuple([str(Pl[0]) for Pl in ResultsT])
                Results = self.cur.execute('''
                    select dvPoint1.PlateNo,dvPoint1.VehType,dVPoint1.DetectionTimeIP,dvPoint2.DetectionTimeIP,dvPoint3.DetectionTimeIP
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2, DetectedVehiclesTemp dvPoint3
                    on 
                    dvPoint1.PlateNo in ''' + str(PlateList) + '''
                    and dvPoint2.PlateNo in ''' + str(PlateList) + '''
                    and dvPoint3.PlateNo in ''' + str(PlateList) + '''
                    and dvPoint1.PlateNo = dvPoint2.PlateNo and dvPoint2.PlateNo = dvPoint3.PlateNo
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ? and dvPoint3.CLCode = ?
                    and dvPoint1.DetectionTimeIP < dvPoint2.DetectionTimeIP and dvPoint2.DetectionTimeIP < dvPoint3.DetectionTimeIP
                    order by dvPoint1.DetectionTime
                    ''', (str(CountLocationNo[0]), str(CountLocationNo[1]), str(CountLocationNo[2]))).fetchall()
                    
                
            else:
                ResultsT = self.cur.execute('''
                    select dvPoint1.PlateNo
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo 
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTime < dvPoint2.DetectionTime
                    intersect
                    select dvPoint1.PlateNo
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo 
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTime < dvPoint2.DetectionTime
                    ''', (str(CountLocationNo[0]), str(CountLocationNo[1]), str(CountLocationNo[1]), str(CountLocationNo[2]))).fetchall()
            
                PlateList = tuple([str(Pl[0]) for Pl in ResultsT])
                Results = self.cur.execute('''
                    select dvPoint1.PlateNo,dvPoint1.VehType,dVPoint1.DetectionTime,dvPoint2.DetectionTime,dvPoint3.DetectionTime
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2, DetectedVehiclesTemp dvPoint3
                    on 
                    dvPoint1.PlateNo in ''' + str(PlateList) + '''
                    and dvPoint2.PlateNo in ''' + str(PlateList) + '''
                    and dvPoint3.PlateNo in ''' + str(PlateList) + '''
                    and dvPoint1.PlateNo = dvPoint2.PlateNo and dvPoint2.PlateNo = dvPoint3.PlateNo
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ? and dvPoint3.CLCode = ?
                    and dvPoint1.DetectionTime < dvPoint2.DetectionTime and dvPoint2.DetectionTime < dvPoint3.DetectionTime
                    order by dvPoint1.DetectionTime
                    ''', (str(CountLocationNo[0]), str(CountLocationNo[1]), str(CountLocationNo[2]))).fetchall()
                    

        if len(CountLocationNo) == 2:
            if self.Interpolate:
                ResultsPL = self.cur.execute('''
                    select distinct dvPoint1.PlateNo
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTimeIP < dvPoint2.DetectionTimeIP 
                    order by dvPoint1.DetectionTimeIP''', (str(CountLocationNo[0]), str(CountLocationNo[1]))).fetchall()
            else:
                ResultsPL = self.cur.execute('''
                    select distinct dvPoint1.PlateNo
                    from DetectedVehiclesTemp dvPoint1 
                    inner join DetectedVehiclesTemp dvPoint2
                    on dvPoint1.PlateNo = dvPoint2.PlateNo
                    and dvPoint1.CLCode = ? and dvPoint2.CLCode = ?
                    and dvPoint1.DetectionTime < dvPoint2.DetectionTime
                    order by dvPoint1.DetectionTime''', (str(CountLocationNo[0]), str(CountLocationNo[1]))).fetchall()
                
                
                
            ResultsPlateList = [Res[0] for Res in Results]
            ResultsDistinctPlates = [Pl[0] for Pl in ResultsPL]
            for Plate in ResultsDistinctPlates:
                PlTempList = []
                Count = ResultsPlateList.count(Plate)
                if Count <= 1:
                    pass
                else:
                    for i in range(Count - 1):
                        ind = ResultsPlateList.index(Plate)
                        item = Results.pop(ind)
                        it = ResultsPlateList.pop(ind)
                    
        self.cur.execute('drop table DetectedVehiclesTemp')
        return Results
    
    def Execute_Flow_Bundle(self, CLs):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Executes flowbundle between conjunctive set of CLs
        
        Ussually Run from Get_Visum_Volume
        
        IN: [CL,CL,CL] (IN: CL -> error, IN:[CL] ok
        
        OUT: Visum Flow Bundle Object
        """

        Segment = self.Visum.Net.DemandSegments.ItemByKey(self.DSeg) #TSys=DSeg????
        FlowBundle = Segment.FlowBundle
        ActivityTypeSet = FlowBundle.CreateActivityTypeSet()
        for CL in CLs:
            FlowBundle.CreateCondition(self.Visum.Net.Links.ItemByKey(CL[3], CL[4]), ActivityTypeSet)
        FlowBundle.ExecuteCurrentConditions()
        return FlowBundle
    
    def Get_Visum_Volume(self, CLs=0):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Gets Visum Volume from Execute_Flow_Bundle
        IN: [CL,CL,CL] (IN: CL -> error, IN:[CL] ok
        
        OUT: Volume of Flow Bundle
        """
        self.Execute_Flow_Bundle(CLs)
        List = self.Visum.Lists.CreateLinkList
        List.AddColumn("VolFlowBundle_TSys(" + self.TSys + ")") #dodaj interesujacy parametr
        return List.Max(0)
    
    def Get_CL_Zones(self, CL):
        """
        TO BE EXPLOITED IN FUTURE:
        
        Procedure telling which zones are using certain CL. Origin Zones + Dest Zones
        """
        self.FlowMatrixPath = "C:\m1.fma"
        FlowBundle = self.Execute_Flow_Bundle([CL])
        
        FlowBundle.Save(self.FlowMatrixPath, 'b') 
        MtxEditor = self.Visum.MatrixEditor  
        MtxEditor.MLoad(self.FlowMatrixPath)
        size = MtxEditor.MGetRowCount()
        Sums = []
        OriginZones = []
        DestZones = []
        Zones = self.Visum.Net.Zones.GetMultiAttValues("No")
        for zone in Zones:
            Sums.append([MtxEditor.MGetOriginSumByIndex(zone[1]), MtxEditor.MGetDestinationSumByIndex(zone[1])])
        for i, sum in enumerate(Sums):
            if sum[0] > 0:
                OriginZones.append([Zones[i][1], sum[0]])    
            if sum[1] > 0:
                DestZones.append([Zones[i][1], sum[1]])
        return OriginZones, DestZones

    def Licz_Zaleznosci_Miedzy_Rejonami(self):
        def contained(A, B):
            if len(A) == 0:
                return False
            if B.count(A[0]) == 0:
                return False
            elif B[B.index(A[0]):B.index(A[0]) + len(A)] == A:
                return True
            else:
                return False
        
        ID = self.cur.execute("SELECT IDD FROM MATRIX").fetchall()
        Pathnodes = self.cur.execute("SELECT PATHNODES FROM MATRIX").fetchall()
        State = self.cur.execute("SELECT STATE FROM MATRIX").fetchall()
        if Pathnodes[0][0] == 'None':
            return
        IDs = []
        for id in ID:
            IDs.append(id[0]) 
        m = len(Pathnodes)
        self.dialog = wx.ProgressDialog ('Progress', "Calculating Matrix Topology", maximum=3 * m + 4)
        n = []
        for Pathnode in Pathnodes:
            if Pathnode != ('[]',):
                P = Pathnode[0]
                P = tuple(float(i) for i in P[1:-1].split(', '))
                n.append([int(i) for i in P])
            else:
                n.append([])
        Pathnodes = n
        Result = []
        
        for s, PN1 in enumerate(Pathnodes):
            self.dialog.Update(s)
            if State[s][0][-4:] in ["loop", "path"]:
                r = 'None'
            else:
                r = []
                for i, PN2 in enumerate(Pathnodes):
                    if PN1 != PN2:
                        if State[i][0][-4:] not in ["loop", "path"]:
                            if contained(PN1, PN2):
                                r.append(i)
            Result.append(r)
        IS_CONTAINED = Result
        Result = []
        for s, PN1 in enumerate(Pathnodes):
            self.dialog.Update(s + m) 
            if State[s][0][-4:] in ["loop", "path"]:
                r = 'None'
            else:
                r = []
                for i, PN2 in enumerate(Pathnodes):                
                    if PN1 != PN2:
                        if State[i][0][-4:] not in ["loop", "path"]:
                            if contained(PN2, PN1):
                                r.append(i)
            Result.append(r)
        CONTAINS = Result
        u = 2 * m
        for i in range(len(IS_CONTAINED)):
            u += 1
            self.dialog.Update(u)
            Query = "UPDATE Matrix SET IS_CONTAINED_IN_IDD = '" + str(IS_CONTAINED[i]) + "', CONTAINS_IDD  = '" + str(CONTAINS[i]) + "' where IDD='" + str(IDs[i]) + "'"               
            self.cur.execute(Query) 
        self.con.commit()
        self.dialog.Destroy()
         
        #print "ids=",IDs[:10] 
        #print "pathnodes=",Pathnodes[:10] 
        #print "cont=",CONTAINS[:10] 
        #print "is_cont=",IS_CONTAINED[:10] 

    def Licz_Nowe_Volumes(self):
        IS_CONTAINED_IN = self.cur.execute("SELECT IS_CONTAINED_IN_IDD FROM MATRIX").fetchall()
        n = []
        for I in IS_CONTAINED_IN:
            if I[0] == 'None':
                n.append(0)
            elif I == ('[]',):
                n.append([])
            else:
                P = I[0]
                P = tuple(float(i) for i in P[1:-1].split(', '))
                n.append([int(i) for i in P])
        IS_CONTAINED_IN = n        
        
        Volume_APNR = self.cur.execute("SELECT APNR_VOLUME_OD FROM MATRIX").fetchall()
        
        #Volume_APNR=self.cur.execute("SELECT VOLUME_VISUM FROM MATRIX").fetchall()
        
        V = []
        for Vol in Volume_APNR:
            if Vol[0] == 'None':
                V.append(0)
            elif Vol[0] == None:
                V.append(0)
            else:
                V.append(int(Vol[0]))
            
        
        Volume_APNR = V
        IDDs = range(len(Volume_APNR))
        New_Volumes = [-1 for i in IDDs]
        
        #pierwsza petla - dodaj niezalezne, nie zawarte w zadnym innym ich NEW_VOL=VOLAPNR
        for i in IDDs:
            if IS_CONTAINED_IN[i] == []:
                New_Volumes[i] = Volume_APNR[i]
                
#        print IS_CONTAINED_IN[-10:]
#        print V[-10:]
#        print New_Volumes[-10:]            
        #petla glowna - dodaj te, dla ktorych wszystkie w ktorych jest zawarta sa okreslone
        for duzy in range(20): 
            for i in IDDs:
                if New_Volumes[i] == -1:
                    gotya = True
                    s = 0
                    if IS_CONTAINED_IN[i] != 0:
                        
                        for p in IS_CONTAINED_IN[i]:
                            #if not gotya:
                            #    break
                            if New_Volumes[p] == -1:
                                gotya = False
                            else:
                               #try: 
                                s += New_Volumes[p] 
                               #except:
                               #    pass                           
                        if gotya:
                            try:
                                New_Volumes[i] = Volume_APNR[i] - s
                            except:
                                New_Volumes[i] = 0
            
            bb = 0
            for nw in New_Volumes:
                if nw == -1:
                   bb += 1
            print bb
            
        for i in IDDs:            
            Query = "UPDATE Matrix SET APNR_VOLUME_DETECTED = '" + str(New_Volumes[i]) + "' where IDD='" + str(i + 1) + "'"               
            self.cur.execute(Query) 
        self.con.commit()

    def Fratar(self):  
        # TO DO: importowac VOL_FRATAR Z VISUMA I UPDATE VOL_FRATAR
        
        
        def suma_wierszy(macierz):           
            """
            zwraca sume wierszy z macierzy
            in list(list)
            return list
            """
            return [sum(wiersz) for wiersz in macierz]
        
        def suma_kolumn(macierz):
            """
            zwraca sume kolumn z macierzy - korzysta z funkcji transpozycja
            in list(list))
            return list
            """
            return suma_wierszy(transpozycja(macierz))
        
        def transpozycja(macierz):
            """
            transponuje macierz 
            in macierz (list(list))
            return macierz
            """
            macierz=[macierz[j][i] for i in range(len(macierz)) for j in range(len(macierz))]
            macierz=self.List2Matrix(macierz)
            return macierz
            
        #Get Params
        
        Fratar_Data=self.cur.execute("SELECT VOL_FRATAR_FROM,VOL_FRATAR_TO,NO FROM COUNTLOCATIONS").fetchall()        
        macierz=self.cur.execute("SELECT APNR_VOLUME_OD FROM MATRIX").fetchall()
        macierz=[el[0] for el in macierz] 
        Vol_Fratar_From=[el[0] for el in Fratar_Data]
        Vol_Fratar_To=[el[1] for el in Fratar_Data]
        ID=[el[2] for el in Fratar_Data]        
             
        #Zrob macierze
        macierz=self.List2Matrix(macierz)        
        
        sumy_wierszy=suma_wierszy(macierz)
        sumy_kolumn=suma_kolumn(macierz)
        
        #Oblicz wspolczynniki wzrostu
        
        wspolczynniki_wierszy=[float(Vol_Fratar_From[i])/float(sumy_wierszy[i]) for i in range(len(Vol_Fratar_From))]
        wspolczynniki_kolumn=[float(Vol_Fratar_To[i])/float(sumy_kolumn[i]) for i in range(len(Vol_Fratar_To))]      
        sredni_wzrost=float((sum(Vol_Fratar_From)+sum(Vol_Fratar_To)))/(2*float(sum(sumy_kolumn)))
#        print sum(Vol_Fratar_From)
#        print sum(Vol_Fratar_To)
#        print sum(sumy_kolumn)
#        print sredni_wzrost
        
        #Wlasciwa procedura Fratar
        nowa_macierz=[[komorka*wspolczynniki_wierszy[i]*wspolczynniki_kolumn[j]/sredni_wzrost 
                      for j,komorka in enumerate(wiersz)] 
                      for i,wiersz in enumerate(macierz)]        
        for i,wiersz in enumerate(nowa_macierz):
            for j,komorka in enumerate(wiersz):
                k=str(komorka)[:(str(komorka).index(".")+3)]                
                self.cur.execute('update Matrix set APNR_VOLUME_FRATAR =? where FromCLCode =? and ToCLCode= ?',(k,ID[i],ID[j]))
        self.con.commit()    
        
    def List2Matrix(self,A):
        size=int(sqrt(len(A)))
        return [A[size*(i):size*(i+1)] for i in range(size)]
    
    def Matrix2List(self,A):
        A=[]
        for row in A:
            for el in row:
                A.append(el)
        return A
           
class Query_Container:
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
    ####
    Nie uzywana klasa do liczenia statystyk filtra    
    """
    def __init__(self, Query_Result, precalc=None):
        self.Filter = Query_Result
        self.precalc = False
        if precalc: 
            self.precalc = True
            self.__precalc()
        
    def __precalc(self):
        self.Get_DetectionFrequencies()
        self.Get_DetectionTimes()
        self.Get_Total()
        self.Get_TimeSpan()
        self.CountUnread = self.Get_CountUnread()
        self.ShareUnread = self.CountUnread / float(self.Total)
        self.Shares = [self.Get_Share_VehType(typ) for typ in ["Car", "Bus", "LKW"]]
    
    def Get(self):
        return self.Filter
        
    def Get_Total(self):
        self.Total = len(self.Filter)
        
    def Get_DetectionTimes(self):
        self.DetectionTimes = [count[2] for count in self.Filter]
   
    def Get_DetectionFrequencies(self):
        ### nie uporzadkowane!!!!!
        self.DetectionFrequencies = [self.Filter[i + 1][2] - self.Filter[i][2] for i, a in enumerate(self.Filter[:-1])]
        
    
    def Get_TimeSpan(self):
        times = self.Get_DetectionTimes()
        self.TimeSpan = max(times) - min(times)
    
    def Get_Share_VehType(self, VehType):
        return len([count for count in self.Filter if count[3] == VehType]) / float(self.Get_Total())
 
    def Get_CountUnread(self, unreadablesign="#"):
        return len([count for count in self.Filter if unreadablesign in count[4]])

    
        
    def Gen_Plot_Points(self, ys):
        return [range(len(ys)), ys]

class PlotPanel (wx.Panel):
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
    ####
    GUI Matplotlib object. 
    Overrides wx.Panel, creates Figure,Canvas, etc.
    Works...  
      
    The PlotPanel has a Figure and a Canvas. OnSize events simply set a 
    flag, and the actual resizing of the figure is triggered by an Idle event."""
    def __init__(self, parent, color=None, dpi=None, **kwargs):
        from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg
        from matplotlib.figure import Figure

        # initialize Panel
        if 'id' not in kwargs.keys():
            kwargs['id'] = wx.ID_ANY
        if 'style' not in kwargs.keys():
            kwargs['style'] = wx.NO_FULL_REPAINT_ON_RESIZE
        wx.Panel.__init__(self, parent, **kwargs)
        self.parent = parent
        # initialize matplotlib stuff
        self.figure = Figure(None, dpi)
        self.canvas = FigureCanvasWxAgg(self, -1, self.figure)
        self._SetSize()
        self.draw()

        self._resizeflag = False

        self.Bind(wx.EVT_IDLE, self._onIdle)
        self.Bind(wx.EVT_SIZE, self._onSize)

    def _onSize(self, event):
        self._resizeflag = True

    def _onIdle(self, evt):
        if self._resizeflag:
            self._resizeflag = False
            self._SetSize()

    def _SetSize(self):
        pixels = tuple(self.parent.GetClientSize())
        self.SetSize(pixels)
        self.canvas.SetSize(pixels)
        self.figure.set_size_inches(float(pixels[0]) / self.figure.get_dpi(),
                                     float(pixels[1]) / self.figure.get_dpi())

    def draw(self): pass # abstract, to be overridden by child classes

class APNR_GUI(wx.Frame):
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski intelligent-infrastructure.eu
    ####
    Main GUI Class. Made with wx glade.
    Init overloads wx.frame with Visum as additional param
    """

    def __init__(self, Visum, *args, **kwds):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski intelligent-infrastructure.eu
        ####
        wx GUI init constructor
        """
        self.Visum = Visum        
        self.Make_File_Paths() 
        # begin wxGlade: APNR_GUI.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, None, -1)
        #wx.Frame.__init__(self, *args, **kwds)
        self.panele = wx.Notebook(self, -1, style=0)
        self.panel_CLs_Matrix = wx.Panel(self.panele, -1)
        self.panel_CLs_Matrix.SetBackgroundColour(wx.Colour(240, 240, 240))
        #self.panele_CLStats = wx.Panel(self.panele, -1)
        #self.panele_CLStats.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.panel_Plot = wx.Panel(self.panele, -1)
        self.panel_Plot.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.panel_Filter = wx.Panel(self.panele, -1)
        self.panel_Filter.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.panel_Init = wx.Panel(self.panele, -1)
        self.panel_Init.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.sizer_7_staticbox = wx.StaticBox(self.panel_Init, -1, "Console")
        self.CLs_sizer_staticbox = wx.StaticBox(self.panel_Init, -1, "")
        self.filter_CL_sizer_staticbox = wx.StaticBox(self.panel_Filter, -1, "count location #1")
        self.filter_CL_sizer2_staticbox = wx.StaticBox(self.panel_Filter, -1, "count location #2")
        self.filter_CL_sizer3_staticbox = wx.StaticBox(self.panel_Filter, -1, "count location #3")
        self.filter_VehType_sizer_staticbox = wx.StaticBox(self.panel_Filter, -1, "vehicle types")
        self.from_time_sizer_staticbox = wx.StaticBox(self.panel_Filter, -1, "from time")
        self.filter_CL_copy_copy_staticbox = wx.StaticBox(self.panel_Filter, -1, "plate no")
        self.to_time_sizer_staticbox = wx.StaticBox(self.panel_Filter, -1, "to time")
        self.Conditions_staticbox = wx.StaticBox(self.panel_Filter, -1, "Filter conditions")
        self.query_results_sizer_staticbox = wx.StaticBox(self.panel_Filter, -1, "Query result")
        self.statistics_sizer_staticbox = wx.StaticBox(self.panel_Filter, -1, "Statistics")
        self.plot_params_sizer_staticbox = wx.StaticBox(self.panel_Plot, -1, "Parameters")
        self.plot_sizer_2_staticbox = wx.StaticBox(self.panel_Plot, -1, "Plot")
        '''
        self.filter_CL_sizer_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "count locations")
        self.filter_VehType_sizer_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "vehicle types")
        self.from_time_sizer_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "from time")
        self.filter_CL_copy_copy_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "plate no")
        self.to_time_sizer_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "to time")
        self.Conditions_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "Filter conditions")
        self.query_results_sizer_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "Query result")
        self.statistics_sizer_CLS_staticbox = wx.StaticBox(self.panele_CLStats, -1, "Statistics")
        '''
        
        
        self.filter_CL_sizer_Mat_staticbox = wx.StaticBox(self.panel_CLs_Matrix, -1, "show matrix values for")
        self.Conditions_Mat_staticbox = wx.StaticBox(self.panel_CLs_Matrix, -1, "Filter conditions")
        self.query_results_sizer_Mat_staticbox = wx.StaticBox(self.panel_CLs_Matrix, -1, "Query result")
        self.statistics_sizer_Mat_staticbox = wx.StaticBox(self.panel_CLs_Matrix, -1, "Statistics")
        self.sizer_8_staticbox = wx.StaticBox(self.panel_Init, -1, "Parameters")
        
        
        self.panel_Paths = wx.Panel(self.panele, -1)
        self.panel_Paths.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.filter_CL_paths = wx.ListBox(self.panel_Paths, -1, choices=[], style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.CL_sizer_paths_staticbox = wx.StaticBox(self.panel_Paths, -1, "from CL")
        self.filter_VehTypes_paths = wx.ListBox(self.panel_Paths, -1, choices=[], style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.filter_VehType_sizer_paths_staticbox = wx.StaticBox(self.panel_Paths, -1, "to CL")
        self.filter_blank_panel_paths = wx.Panel(self.panel_Paths, -1)
        self.btn_import = wx.Button(self.filter_blank_panel_paths, -1, "commit changes")
        self.btn_filter_copy = wx.Button(self.filter_blank_panel_paths, -1, "filter/update")
        self.panel_p = wx.Panel(self.panel_Paths, -1)
        self.paths_staticbox = wx.StaticBox(self.panel_Paths, -1, "Filter conditions")
        self.grid_1_paths = wx.grid.Grid(self.panel_Paths, -1, size=(1, 1))
        
        
        self.query_results_sizer_paths_staticbox = wx.StaticBox(self.panel_Paths, -1, "Query result")
        self.statistics_sizer_paths_staticbox = wx.StaticBox(self.panel_Paths, -1, "Statistics")
        
        
        # Menu Bar
        self.APNR_menubar = wx.MenuBar()
        wxglade_tmp_menu = wx.Menu()
        self.MenuS_I = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Initialize database", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.MenuS_I)
        self.MenuS_C = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Connect with database", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.MenuS_C)
        self.MenuS_Im = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Import Results", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.MenuS_Im)
        self.MenuS_Pr = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Process Database", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.MenuS_Pr)
        self.MenuS_Fr = wx.MenuItem(wxglade_tmp_menu, wx.NewId(), "Extrapolate Fratar", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu.AppendItem(self.MenuS_Fr)
        wxglade_tmp_menu_sub = wx.Menu()
        wxglade_tmp_menu_sub3 = wx.Menu()
        wxglade_tmp_menu_sub_sub = wx.Menu()
        wxglade_tmp_menu_sub_sub2 = wx.Menu()        
        self.MenuE_E_F = wx.MenuItem(wxglade_tmp_menu_sub_sub, wx.NewId(), "Filter results", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu_sub_sub.AppendItem(self.MenuE_E_F)
        self.Menu_E_E_C = wx.MenuItem(wxglade_tmp_menu_sub_sub, wx.NewId(), "CL statistics", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu_sub_sub.AppendItem(self.Menu_E_E_C)
        self.Menu_E_E_M = wx.MenuItem(wxglade_tmp_menu_sub_sub, wx.NewId(), "Matrix", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu_sub_sub.AppendItem(self.Menu_E_E_M)
        self.MenuE_E_S = wx.MenuItem(wxglade_tmp_menu_sub_sub, wx.NewId(), "Statistics", "", wx.ITEM_NORMAL)
        wxglade_tmp_menu_sub_sub.AppendItem(self.MenuE_E_S)
        
        wxglade_tmp_menu_sub.AppendMenu(wx.NewId(), "to Excel", wxglade_tmp_menu_sub_sub, "")
        
        self.MenuE_V_Z = wx.MenuItem(wxglade_tmp_menu_sub_sub2, wx.NewId(), "Create Zones", "", wx.ITEM_NORMAL)
        self.MenuE_V_P = wx.MenuItem(wxglade_tmp_menu_sub_sub2, wx.NewId(), "Paths", "", wx.ITEM_NORMAL)
        self.MenuE_V_M = wx.MenuItem(wxglade_tmp_menu_sub_sub2, wx.NewId(), "Matrix", "", wx.ITEM_NORMAL)
         
        
        wxglade_tmp_menu_sub_sub2.AppendItem(self.MenuE_V_Z)
        wxglade_tmp_menu_sub_sub2.AppendItem(self.MenuE_V_P)
        wxglade_tmp_menu_sub_sub2.AppendItem(self.MenuE_V_M)
        
        self.MenuI_min = wx.MenuItem(wxglade_tmp_menu_sub3, wx.NewId(), "Import SkimMtx of min travel times", "", wx.ITEM_NORMAL)
        self.MenuI_max = wx.MenuItem(wxglade_tmp_menu_sub3, wx.NewId(), "Import SkimMtx of max travel times", "", wx.ITEM_NORMAL) 
        
        wxglade_tmp_menu_sub3.AppendItem(self.MenuI_min)
        wxglade_tmp_menu_sub3.AppendItem(self.MenuI_max)
        
        wxglade_tmp_menu_sub.AppendMenu(wx.NewId(), "to Visum", wxglade_tmp_menu_sub_sub2, "")
        wxglade_tmp_menu.AppendMenu(wx.NewId(), "Export", wxglade_tmp_menu_sub, "")
        wxglade_tmp_menu.AppendMenu(wx.NewId(), "Import", wxglade_tmp_menu_sub3, "")
        self.APNR_menubar.Append(wxglade_tmp_menu, "Menu")
        self.SetMenuBar(self.APNR_menubar)
        # Menu Bar end
        self.header_title = wx.StaticText(self, -1, "Plate Number Recognition Support by:")
        self.logo = wx.StaticBitmap(self, -1, wx.Bitmap(self.Paths["Logo"], wx.BITMAP_TYPE_ANY))
        self.txt1 = wx.StaticText(self.panel_Init, -1, "1. Demand Segment")
        self.DSeg_Combo = wx.ComboBox(self.panel_Init, -1, choices=[], style=wx.CB_DROPDOWN)
        self.txt2 = wx.StaticText(self.panel_Init, -1, "2. Transport Systems")
        self.TSys_Combo = wx.ComboBox(self.panel_Init, -1, choices=[], style=wx.CB_DROPDOWN)
        self.txt3 = wx.StaticText(self.panel_Init, -1, "3. Interpolate detection times?")
        self.Interpolate_CBox = wx.CheckBox(self.panel_Init, -1, "")
        self.Console = wx.TextCtrl(self.panel_Init, -1, "", style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_CENTRE)
        self.grid_init = wx.grid.Grid(self.panel_Init, -1, size=(1, 1))
        self.list_CL_filter = wx.ListBox(self.panel_Filter, -1, choices=[], style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.list_CL_filter2 = wx.ListBox(self.panel_Filter, -1, choices=[], style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.list_CL_filter3 = wx.ListBox(self.panel_Filter, -1, choices=[], style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.list_VehTypes_filter = wx.ListBox(self.panel_Filter, -1, choices=[], style=wx.LB_MULTIPLE | wx.LB_NEEDED_SB)
        self.from_time_text = wx.TextCtrl(self.panel_Filter, -1, "06:00:00")
        self.plate_no_text = wx.TextCtrl(self.panel_Filter, -1, "None")
        self.to_time_text = wx.TextCtrl(self.panel_Filter, -1, "07:00:00")
        self.btn_filter = wx.Button(self.panel_Filter, -1, "filter")
        self.filter_blank_panel = wx.Panel(self.panel_Filter, -1)
        self.grid_1 = wx.grid.Grid(self.panel_Filter, -1, size=(1, 1))
        self.grid_stats = wx.grid.Grid(self.panel_Filter, -1, size=(1, 1))
        self.combo_box_1 = wx.ComboBox(self.panel_Plot, -1, choices=["Detection Time agains time for single CL", "Histogram of detections against time for single CL", "Travel time against time for pair of CLs", "Histogram of travel times for pair of CLs"], style=wx.CB_DROPDOWN)
        self.combo_box_1.SetSelection(0)
        self.btn_plot = wx.Button(self.panel_Plot, -1, "Plot")
        self.btn_export_excel = wx.Button(self.panel_Plot, -1, "Export to excel")
        self.plot_support_panel = wx.Panel(self.panel_Plot, -1)
        '''
        self.list_CL_filter_CLS = wx.ListBox(self.panele_CLStats, -1, choices=[], style=wx.LB_MULTIPLE|wx.LB_NEEDED_SB)
        self.list_VehTypes_filter_CLS = wx.ListBox(self.panele_CLStats, -1, choices=[], style=wx.LB_MULTIPLE|wx.LB_NEEDED_SB)
        self.from_time_text_CLS = wx.TextCtrl(self.panele_CLStats, -1, "25200")
        self.plate_no_text_CLS = wx.TextCtrl(self.panele_CLStats, -1, "None")
        self.to_time_text_CLS = wx.TextCtrl(self.panele_CLStats, -1, "28800")
        self.btn_filter_CLS = wx.Button(self.panele_CLStats, -1, "filter")
        self.filter_blank_panel_CLS = wx.Panel(self.panele_CLStats, -1)
        self.grid_1_CLS = wx.grid.Grid(self.panele_CLStats, -1, size=(1, 1))
        '''
        self.list_CL_filter_Mat = wx.ListBox(self.panel_CLs_Matrix, -1, choices=[], style=wx.LB_SINGLE | wx.LB_NEEDED_SB)
        self.btn_filter_Mat_copy = wx.Button(self.panel_CLs_Matrix, -1, "Calc Values")
        self.btn_filter_Mat = wx.Button(self.panel_CLs_Matrix, -1, "Show values")
        self.btn_Export_Paths_2Visum = wx.Button(self.panel_CLs_Matrix, -1, "Export Paths to Visum")
        self.panel_1 = wx.Panel(self.panel_CLs_Matrix, -1)
        self.grid_1_Mat = wx.grid.Grid(self.panel_CLs_Matrix, -1, size=(1, 1))
        self.HelpBtn = wx.Button(self, -1, "Help")
        self.panel_2 = wx.Panel(self, -1)
        self.CancelBtn = wx.Button(self, -1, "Cancel")
        
        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_MENU, self.__handler_DB_init, self.MenuS_I)
        self.Bind(wx.EVT_MENU, self.__handler_DB_connect, self.MenuS_C)
        self.Bind(wx.EVT_MENU, self.__handler_import, self.MenuS_Im)
        self.Bind(wx.EVT_MENU, self.__handler_process_menu, self.MenuS_Pr)
        self.Bind(wx.EVT_MENU, self.__handler_fratar, self.MenuS_Fr)
        self.Bind(wx.EVT_MENU, self.__handler_Export_Filter, self.MenuE_E_F)
        self.Bind(wx.EVT_MENU, self.__handler_Export_CL, self.Menu_E_E_C)
        self.Bind(wx.EVT_MENU, self.__handler_export_matrix, self.Menu_E_E_M)
        self.Bind(wx.EVT_MENU, self.__handler_Export_Visum_Zones, self.MenuE_V_Z)
        self.Bind(wx.EVT_MENU, self.__handler_Export_Paths_2_Visum, self.MenuE_V_P)
        self.Bind(wx.EVT_MENU, self.__handler_export_Visum_Matrix, self.MenuE_V_M)
        self.Bind(wx.EVT_MENU, self.__handler_export_Statistics, self.MenuE_E_S)
        self.Bind(wx.EVT_MENU, self.__handler_Import_Skim_Min, self.MenuI_min)
        self.Bind(wx.EVT_MENU, self.__handler_Import_Skim_Max, self.MenuI_max)
        
        self.Bind(wx.EVT_COMBOBOX, self.Update_DSeg, self.DSeg_Combo)
        self.Bind(wx.EVT_COMBOBOX, self.Update_TSys, self.TSys_Combo)
        self.Bind(wx.EVT_CHECKBOX, self.Update_Interpolate, self.Interpolate_CBox)

        self.Bind(wx.grid.EVT_GRID_CMD_CELL_LEFT_CLICK, self.__handler_CLs_click, self.grid_init)
        self.Bind(wx.grid.EVT_GRID_CMD_CELL_LEFT_CLICK, self.__handler_Mtx_click, self.grid_1_Mat)
        self.Bind(wx.grid.EVT_GRID_CMD_CELL_LEFT_DCLICK, self.__handler_Path_click, self.grid_1_paths)
        
        self.Bind(wx.EVT_BUTTON, self.__handler_filter, self.btn_filter)
        self.Bind(wx.EVT_BUTTON, self.__handler_excel_plot_export, self.btn_export_excel)
        self.Bind(wx.EVT_BUTTON, self.__handler_GUI_Plot, self.btn_plot)
        #self.Bind(wx.EVT_BUTTON, self.__handler_filter_ST, self.btn_filter_CLS)
        self.Bind(wx.EVT_BUTTON, self.__handler_calc_matrix, self.btn_filter_Mat_copy)
        self.Bind(wx.EVT_BUTTON, self.__handler_fill_matrix, self.btn_filter_Mat)
        self.Bind(wx.EVT_BUTTON, self.__handler_Export_Paths_2_Visum, self.btn_Export_Paths_2Visum)
        self.Bind(wx.EVT_BUTTON, self.__handler_help, self.HelpBtn)
        self.Bind(wx.EVT_BUTTON, self.__handler_cancel_click, self.CancelBtn)
        
        self.Bind(wx.EVT_BUTTON, self.__handler_savePthtoDb, self.btn_import)
        self.Bind(wx.EVT_BUTTON, self.handler_filtrujPth, self.btn_filter_copy)
        # end wxGlade
        self.__init_Console()
        self.__init_DSeg_TSys()
        self.Interpolate = True
        self.Interpolate_CBox.SetValue(True)
        self.Importer="PK"
        
    def __set_properties(self):
        # begin wxGlade: APNR_GUI.__set_properties
        self.SetTitle("APNR support by i2")
        self.SetSize((991, 890))
        self.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.header_title.SetMinSize((-1, 16))
        self.header_title.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.logo.SetMinSize((-1, 20))
        self.logo.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.txt1.SetMinSize((-1, 16))
        self.txt1.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.txt2.SetMinSize((-1, 16))
        self.txt2.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.txt3.SetMinSize((-1, 16))
        self.txt3.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.Console.SetMinSize((100, -1))
        self.Console.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.grid_init.CreateGrid(1, 1)
        self.grid_init.SetColLabelValue(0, "")
        self.btn_filter.SetMinSize((100, -1))
        self.grid_1.CreateGrid(1, 0)
        self.btn_plot.SetMinSize((87, -1))
        self.btn_export_excel.SetMinSize((87, -1))
        #self.btn_filter_CLS.SetMinSize((100, -1))
        #self.grid_1_CLS.CreateGrid(1, 1)
        #self.grid_1_CLS.SetColLabelValue(0, "")
        self.btn_filter_Mat_copy.SetMinSize((150, -1))
        self.btn_filter_Mat.SetMinSize((150, -1))
        self.btn_Export_Paths_2Visum.SetMinSize((150, -1))
        self.grid_1_Mat.CreateGrid(1, 1)
        self.grid_stats.CreateGrid(1, 7)
        self.grid_stats.SetRowLabelSize(1)
        self.grid_stats.SetColLabelValue(0, "Vol")
        self.grid_stats.SetColLabelValue(1, "Error")
        self.grid_stats.SetColLabelValue(2, "T_MIN")
        self.grid_stats.SetColLabelValue(3, "T_MEAN")
        self.grid_stats.SetColLabelValue(4, "T_MOD")
        self.grid_stats.SetColLabelValue(5, "T_MAX")
        
        self.btn_import.SetMinSize((100, -1))
        
        self.btn_filter_copy.SetMinSize((100, -1))
        
        self.grid_1_paths.CreateGrid(1, 7)
        self.grid_1_paths.SetRowLabelSize(1)
        
        self.grid_1_paths.SetColLabelValue(0, "FromCL")
        self.grid_1_paths.SetColLabelValue(1, "ToCL")
        self.grid_1_paths.SetColLabelValue(2, "enabled")
        self.grid_1_paths.SetColLabelValue(3, "t0")
        self.grid_1_paths.SetColLabelValue(4, "tCur")
        self.grid_1_paths.SetColLabelValue(5, "min t")
        self.grid_1_paths.SetColLabelValue(6, "max t")
        
        self.grid_1_paths.EnableEditing(True)
        
        self.grid_1_Mat.SetColLabelValue(0, "")
        self.HelpBtn.SetMinSize((100, -1))
        self.CancelBtn.SetMinSize((100, -1))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: APNR_GUI.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        Stopka = wx.BoxSizer(wx.HORIZONTAL)
        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        FIlter_sizer_Mat = wx.BoxSizer(wx.VERTICAL)
        sizer_9_Mat = wx.BoxSizer(wx.VERTICAL)
        statistics_sizer_Mat = wx.StaticBoxSizer(self.statistics_sizer_Mat_staticbox, wx.HORIZONTAL)
        query_results_sizer_Mat = wx.StaticBoxSizer(self.query_results_sizer_Mat_staticbox, wx.HORIZONTAL)
        Conditions_Mat = wx.StaticBoxSizer(self.Conditions_Mat_staticbox, wx.HORIZONTAL)
        sizer_3 = wx.BoxSizer(wx.VERTICAL)
        filter_CL_sizer_Mat = wx.StaticBoxSizer(self.filter_CL_sizer_Mat_staticbox, wx.HORIZONTAL)
        #FIlter_sizer_CLS = wx.BoxSizer(wx.VERTICAL)
        sizer_9_CLS = wx.BoxSizer(wx.VERTICAL)
        #statistics_sizer_CLS = wx.StaticBoxSizer(self.statistics_sizer_CLS_staticbox, wx.HORIZONTAL)
        #query_results_sizer_CLS = wx.StaticBoxSizer(self.query_results_sizer_CLS_staticbox, wx.HORIZONTAL)
        #Conditions_CLS = wx.StaticBoxSizer(self.Conditions_CLS_staticbox, wx.HORIZONTAL)
        #fourth_filter_sizer_CLS = wx.BoxSizer(wx.VERTICAL)
        #filter_btn_sizer_CLS = wx.BoxSizer(wx.HORIZONTAL)
        #to_time_sizer_CLS = wx.StaticBoxSizer(self.to_time_sizer_CLS_staticbox, wx.HORIZONTAL)
        #third_filter_sizer_CLS = wx.BoxSizer(wx.VERTICAL)
        #filter_CL_copy_copy_CLS = wx.StaticBoxSizer(self.filter_CL_copy_copy_CLS_staticbox, wx.HORIZONTAL)
        #from_time_sizer_CLS = wx.StaticBoxSizer(self.from_time_sizer_CLS_staticbox, wx.HORIZONTAL)
        #filter_VehType_sizer_CLS = wx.StaticBoxSizer(self.filter_VehType_sizer_CLS_staticbox, wx.HORIZONTAL)
        #filter_CL_sizer_CLS = wx.StaticBoxSizer(self.filter_CL_sizer_CLS_staticbox, wx.HORIZONTAL)
        Plot_sizer = wx.BoxSizer(wx.VERTICAL)
        plot_sizer_2 = wx.StaticBoxSizer(self.plot_sizer_2_staticbox, wx.HORIZONTAL)
        plot_params_sizer = wx.StaticBoxSizer(self.plot_params_sizer_staticbox, wx.HORIZONTAL)
        sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
        FIlter_sizer = wx.BoxSizer(wx.VERTICAL)
        sizer_9 = wx.BoxSizer(wx.VERTICAL)
        statistics_sizer = wx.StaticBoxSizer(self.statistics_sizer_staticbox, wx.HORIZONTAL)
        query_results_sizer = wx.StaticBoxSizer(self.query_results_sizer_staticbox, wx.HORIZONTAL)
        Conditions = wx.StaticBoxSizer(self.Conditions_staticbox, wx.HORIZONTAL)
        fourth_filter_sizer = wx.BoxSizer(wx.VERTICAL)
        filter_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        to_time_sizer = wx.StaticBoxSizer(self.to_time_sizer_staticbox, wx.HORIZONTAL)
        third_filter_sizer = wx.BoxSizer(wx.VERTICAL)
        filter_CL_copy_copy = wx.StaticBoxSizer(self.filter_CL_copy_copy_staticbox, wx.HORIZONTAL)
        from_time_sizer = wx.StaticBoxSizer(self.from_time_sizer_staticbox, wx.HORIZONTAL)
        filter_VehType_sizer = wx.StaticBoxSizer(self.filter_VehType_sizer_staticbox, wx.HORIZONTAL)
        filter_CL_sizer3 = wx.StaticBoxSizer(self.filter_CL_sizer3_staticbox, wx.HORIZONTAL)
        filter_CL_sizer2 = wx.StaticBoxSizer(self.filter_CL_sizer2_staticbox, wx.HORIZONTAL)
        filter_CL_sizer = wx.StaticBoxSizer(self.filter_CL_sizer_staticbox, wx.HORIZONTAL)
        sizer_4 = wx.BoxSizer(wx.VERTICAL)
        CLs_sizer = wx.StaticBoxSizer(self.CLs_sizer_staticbox, wx.HORIZONTAL)
        sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_7 = wx.StaticBoxSizer(self.sizer_7_staticbox, wx.HORIZONTAL)
        sizer_8 = wx.StaticBoxSizer(self.sizer_8_staticbox, wx.HORIZONTAL)
        grid_sizer_1 = wx.GridSizer(3, 2, 0, 0)
        Naglowek = wx.BoxSizer(wx.HORIZONTAL)
        Naglowek.Add(self.header_title, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 10)
        Naglowek.Add(self.logo, 2, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_1.Add(Naglowek, 1, wx.EXPAND, 0)
        grid_sizer_1.Add(self.txt1, 2, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.DSeg_Combo, 3, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 0)
        grid_sizer_1.Add(self.txt2, 2, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.TSys_Combo, 3, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 0)
        grid_sizer_1.Add(self.txt3, 2, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.Interpolate_CBox, 0, wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_8.Add(grid_sizer_1, 1, wx.EXPAND, 0)
        sizer_6.Add(sizer_8, 1, wx.EXPAND, 0)
        sizer_7.Add(self.Console, 2, wx.EXPAND, 0)
        sizer_6.Add(sizer_7, 1, wx.EXPAND, 0)
        sizer_4.Add(sizer_6, 4, wx.EXPAND, 0)
        CLs_sizer.Add(self.grid_init, 1, wx.EXPAND, 0)
        sizer_4.Add(CLs_sizer, 8, wx.EXPAND, 0)
        self.panel_Init.SetSizer(sizer_4)
        filter_CL_sizer.Add(self.list_CL_filter, 1, wx.ALL | wx.EXPAND, 5)
        Conditions.Add(filter_CL_sizer, 1, wx.EXPAND, 0)
        filter_CL_sizer2.Add(self.list_CL_filter2, 1, wx.ALL | wx.EXPAND, 5)
        Conditions.Add(filter_CL_sizer2, 1, wx.EXPAND, 0)
        filter_CL_sizer3.Add(self.list_CL_filter3, 1, wx.ALL | wx.EXPAND, 5)
        Conditions.Add(filter_CL_sizer3, 1, wx.EXPAND, 0)
        filter_VehType_sizer.Add(self.list_VehTypes_filter, 1, wx.ALL | wx.EXPAND, 5)
        Conditions.Add(filter_VehType_sizer, 1, wx.EXPAND, 0)
        from_time_sizer.Add(self.from_time_text, 1, wx.ALL, 5)
        third_filter_sizer.Add(from_time_sizer, 1, wx.EXPAND, 0)
        filter_CL_copy_copy.Add(self.plate_no_text, 1, wx.ALL, 5)
        third_filter_sizer.Add(filter_CL_copy_copy, 1, wx.EXPAND, 0)
        Conditions.Add(third_filter_sizer, 1, wx.EXPAND, 0)
        to_time_sizer.Add(self.to_time_text, 1, wx.ALL, 5)
        fourth_filter_sizer.Add(to_time_sizer, 1, wx.EXPAND, 0)
        filter_btn_sizer.Add(self.btn_filter, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 25)
        fourth_filter_sizer.Add(filter_btn_sizer, 1, wx.EXPAND, 0)
        Conditions.Add(fourth_filter_sizer, 1, wx.EXPAND, 0)
        Conditions.Add(self.filter_blank_panel, 2, wx.EXPAND, 0)
        FIlter_sizer.Add(Conditions, 3, wx.EXPAND, 0)
        query_results_sizer.Add(self.grid_1, 1, wx.EXPAND, 0)
        sizer_9.Add(query_results_sizer, 8, wx.EXPAND, 0)
        statistics_sizer.Add(self.grid_stats, 2, wx.EXPAND, 0)
        sizer_9.Add(statistics_sizer, 2, wx.EXPAND, 0)
        FIlter_sizer.Add(sizer_9, 10, wx.EXPAND, 0)
        self.panel_Filter.SetSizer(FIlter_sizer)
        sizer_5.Add(self.combo_box_1, 3, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_5.Add(self.btn_plot, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 25)
        sizer_5.Add(self.btn_export_excel, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 25)
        plot_params_sizer.Add(sizer_5, 1, wx.EXPAND, 0)
        Plot_sizer.Add(plot_params_sizer, 1, wx.EXPAND, 0)
        plot_sizer_2.Add(self.plot_support_panel, 1, wx.EXPAND, 0)
        Plot_sizer.Add(plot_sizer_2, 6, wx.EXPAND, 0)
        self.panel_Plot.SetSizer(Plot_sizer)
        '''
        filter_CL_sizer_CLS.Add(self.list_CL_filter_CLS, 1, wx.ALL|wx.EXPAND, 5)
        Conditions_CLS.Add(filter_CL_sizer_CLS, 1, wx.EXPAND, 0)
        filter_VehType_sizer_CLS.Add(self.list_VehTypes_filter_CLS, 1, wx.ALL|wx.EXPAND, 5)
        Conditions_CLS.Add(filter_VehType_sizer_CLS, 1, wx.EXPAND, 0)
        from_time_sizer_CLS.Add(self.from_time_text_CLS, 1, wx.ALL, 5)
        third_filter_sizer_CLS.Add(from_time_sizer_CLS, 1, wx.EXPAND, 0)
        filter_CL_copy_copy_CLS.Add(self.plate_no_text_CLS, 1, wx.ALL, 5)
        third_filter_sizer_CLS.Add(filter_CL_copy_copy_CLS, 1, wx.EXPAND, 0)
        Conditions_CLS.Add(third_filter_sizer_CLS, 1, wx.EXPAND, 0)
        to_time_sizer_CLS.Add(self.to_time_text_CLS, 1, wx.ALL, 5)
        fourth_filter_sizer_CLS.Add(to_time_sizer_CLS, 1, wx.EXPAND, 0)
        filter_btn_sizer_CLS.Add(self.btn_filter_CLS, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 25)
        fourth_filter_sizer_CLS.Add(filter_btn_sizer_CLS, 1, wx.EXPAND, 0)
        Conditions_CLS.Add(fourth_filter_sizer_CLS, 1, wx.EXPAND, 0)
        Conditions_CLS.Add(self.filter_blank_panel_CLS, 2, wx.EXPAND, 0)
        FIlter_sizer_CLS.Add(Conditions_CLS, 3, wx.EXPAND, 0)
        query_results_sizer_CLS.Add(self.grid_1_CLS, 1, wx.EXPAND, 0)
        sizer_9_CLS.Add(query_results_sizer_CLS, 8, wx.EXPAND, 0)
        sizer_9_CLS.Add(statistics_sizer_CLS, 1, wx.EXPAND, 0)
        FIlter_sizer_CLS.Add(sizer_9_CLS, 10, wx.EXPAND, 0)
        
        #self.panele_CLStats.SetSizer(FIlter_sizer_CLS)
        '''
        filter_CL_sizer_Mat.Add(self.list_CL_filter_Mat, 1, wx.ALL | wx.EXPAND, 5)
        Conditions_Mat.Add(filter_CL_sizer_Mat, 3, wx.EXPAND, 0)
        sizer_3.Add(self.btn_filter_Mat_copy, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_3.Add(self.btn_filter_Mat, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_3.Add(self.btn_Export_Paths_2Visum, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        Conditions_Mat.Add(sizer_3, 1, wx.EXPAND, 0)
        Conditions_Mat.Add(self.panel_1, 1, wx.EXPAND, 0)
        FIlter_sizer_Mat.Add(Conditions_Mat, 3, wx.EXPAND, 0)
        query_results_sizer_Mat.Add(self.grid_1_Mat, 1, wx.EXPAND, 0)

        sizer_9_Mat.Add(query_results_sizer_Mat, 8, wx.EXPAND, 0)      
        statistics_sizer_Mat.Add(self.grid_stats, 2, wx.EXPAND, 0)  
        sizer_9_Mat.Add(statistics_sizer_Mat, 1, wx.EXPAND, 0)
        FIlter_sizer_Mat.Add(sizer_9_Mat, 10, wx.EXPAND, 0)
        
        self.panel_CLs_Matrix.SetSizer(FIlter_sizer_Mat)
        self.panele.AddPage(self.panel_Init, "Start")
        self.panele.AddPage(self.panel_Filter, "General DB Queries")
        self.panele.AddPage(self.panel_Plot, "Plot")
        #self.panele.AddPage(self.panele_CLStats, "Aggregated statistics - disabled")
        self.panele.AddPage(self.panel_CLs_Matrix, "Matrix")
        sizer_2.Add(self.panele, 1, wx.ALL | wx.EXPAND, 5)
        sizer_1.Add(sizer_2, 20, wx.EXPAND, 0)
        Stopka.Add(self.HelpBtn, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        Stopka.Add(self.panel_2, 4, wx.EXPAND, 0)
        Stopka.Add(self.CancelBtn, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 10)
        #TO DO SN: CZEMU KLEIN NIE WIDZI HELP+CANCEL? Stopka.Add(self.HelpBtn, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        #nie moze nie widziec, naprawde
        
        sizer_1.Add(Stopka, 2, wx.ALL | wx.EXPAND, 0)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade
        
        self.PlotPanel = PlotPanel(self.plot_support_panel)
        self.Figure = self.PlotPanel.figure
        self.Subplot = self.Figure.add_subplot(111)
        
        
        paths_sizer = wx.BoxSizer(wx.VERTICAL)
        sizer_9_paths = wx.BoxSizer(wx.VERTICAL)
        self.statistics_sizer_paths_staticbox.Lower()
        statistics_sizer_paths = wx.StaticBoxSizer(self.statistics_sizer_paths_staticbox, wx.HORIZONTAL)
        self.query_results_sizer_paths_staticbox.Lower()
        query_results_sizer_paths = wx.StaticBoxSizer(self.query_results_sizer_paths_staticbox, wx.HORIZONTAL)
        self.paths_staticbox.Lower()
        paths = wx.StaticBoxSizer(self.paths_staticbox, wx.HORIZONTAL)
        
        self.filter_VehType_sizer_paths_staticbox.Lower()
        filter_VehType_sizer_paths = wx.StaticBoxSizer(self.filter_VehType_sizer_paths_staticbox, wx.HORIZONTAL)
        self.CL_sizer_paths_staticbox.Lower()
        CL_sizer_paths = wx.StaticBoxSizer(self.CL_sizer_paths_staticbox, wx.HORIZONTAL)
        
        CL_sizer_paths.Add(self.filter_CL_paths, 1, wx.ALL | wx.EXPAND, 5)
        paths.Add(CL_sizer_paths, 1, wx.EXPAND, 0)
        filter_VehType_sizer_paths.Add(self.filter_VehTypes_paths, 1, wx.ALL | wx.EXPAND, 5)
        paths.Add(filter_VehType_sizer_paths, 1, wx.EXPAND, 0)
        
        sizer_12 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_12.Add(self.btn_import, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 25)
        
        
        filter_btn_sizer_copy = wx.BoxSizer(wx.HORIZONTAL)
        filter_btn_sizer_copy.Add(self.btn_filter_copy, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 25)
        
        fourth_filter_sizer_copy = wx.BoxSizer(wx.VERTICAL)
        fourth_filter_sizer_copy.Add(sizer_12, 1, wx.EXPAND, 0)
        fourth_filter_sizer_copy.Add(filter_btn_sizer_copy, 1, wx.EXPAND, 0)
        
        #self.filter_blank_panel_paths = wx.Panel(self.panel_Paths, -1)
        self.filter_blank_panel_paths.SetSizer(fourth_filter_sizer_copy)
        paths.Add(self.filter_blank_panel_paths, 1, wx.EXPAND, 0)
        paths.Add(self.panel_p, 2, wx.EXPAND, 0)
        paths_sizer.Add(paths, 3, wx.EXPAND, 0)
        query_results_sizer_paths.Add(self.grid_1_paths, 1, wx.EXPAND, 0)
        sizer_9_paths.Add(query_results_sizer_paths, 8, wx.EXPAND, 0)
        sizer_9_paths.Add(statistics_sizer_paths, 1, wx.EXPAND, 0)
        paths_sizer.Add(sizer_9_paths, 10, wx.EXPAND, 0)
        self.panel_Paths.SetSizer(paths_sizer)
        
        self.panele.AddPage(self.panel_Paths, "Paths")
        #DONE SN: Ladniej zamknac ten panel
    
    def __init_DSeg_TSys(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski intelligent-infrastructure.eu
        ####
        run from __init__
        fill Combo Boxes with DSegments and TSystems
        """
        try:
            Segments = self.Visum.Net.DemandSegments.GetMultiAttValues("Code")
            TSyss = self.Visum.Net.TSystems.GetMultiAttValues("CODE")
        except:
            return
        self.DSeg_Combo.AppendItems([str(Segments[s][1]) for s in range(len(Segments))])
        self.TSys_Combo.AppendItems([str(TSyss[s][1]) for s in range(len(TSyss))])
        
        try:
            self.DSeg_Combo.Select(0)
        except:
            pass
        try:
            self.TSys_Combo.Select(1)
        except:
            pass
        
    def __init_Console(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski intelligent-infrastructure.eu
        ####
        run from __init__
        fill console with initial lines
        """
        try:
            2+2            
        except:
            pass
        #self.ErrMsg("xlrd module for reading excel files not found\ninstall it from Add-On installation folder, or download it.\nxlrd is needed only to import xls data")
        
        #self.Paths["Report"]=self.Paths["ScriptFolder"] + "\\report.txt"
        PK=True
        if not PK:
            filereport = open('report.txt', 'w')

        self.__updateConsole("")
        self.__updateConsole("i2")
        self.__updateConsole("")
        self.__updateConsole("intelligent - infrastructure")
        self.__updateConsole("")
        self.__updateConsole("visum scripts and applications")
        self.__updateConsole("")
        self.__updateConsole("")
        self.__updateConsole("Automatic Plate Number Recognition Support")
        self.__updateConsole("")
        self.__updateConsole("Help can be found in your resource files (help.html)")
        self.__updateConsole("Additional support: info@intelligent-infrastructure.eu")
        self.__updateConsole("")
        self.__updateConsole("")
        self.__updateConsole("You can start with: ")
        self.__updateConsole("a) creating new DataBase and importing counting results")
        self.__updateConsole("b) opening existing database file")
    def __updateConsole(self, flag, t=None): 
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        add line "flag" to console (creates new line).
        Multiline text with Vscroll, focus down.
        
        """
        PK=True
        if not PK:
            flag += "\n"
            self.Console.AppendText(flag)
            self.Console.Refresh()
            #DONE SN WYRZUC REPORT DO PLIKU TXT
            
            filereport= open(self.Paths["Report"], 'a')
            filereport.write(flag)
            filereport.close()
    
    def __populate_filters(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        On Panel there are filters for DB filtering, they should take values after DB connection is established.
        Values are taken from DB: countlocations, VehType... 
        """        
        self.DB.cur.execute("Select Distinct CLCODE from CountLocations")
        CLs = self.DB.cur.fetchall()
        self.__updateConsole("There are " + str(len(CLs)) + " Count Locations imported from Visum")
        CLs = [CL[0] for CL in CLs]
        self.DB.cur.execute("Select Distinct VehType from DetectedVehicles")
        VehTypes = self.DB.cur.fetchall()
        VehTypes = [VehType[0] for VehType in VehTypes]
        
        self.list_CL_filter.SetItems([])
        self.list_CL_filter.Append("Any")
        self.list_CL_filter.AppendItems(CLs)
        self.list_CL_filter.Select(0)
        
        self.list_CL_filter2.SetItems([])
        self.list_CL_filter2.Append("Any")
        self.list_CL_filter2.AppendItems(CLs)
        self.list_CL_filter2.Select(0)
        
        self.list_CL_filter3.SetItems([])
        self.list_CL_filter3.Append("Any")
        self.list_CL_filter3.AppendItems(CLs)
        self.list_CL_filter3.Select(0)
        
        self.list_VehTypes_filter.SetItems([])
        self.list_VehTypes_filter.Append("Any")
        self.list_VehTypes_filter.AppendItems(VehTypes)
        self.list_VehTypes_filter.Select(0)
        
        self.filter_CL_paths.SetItems([])
        self.filter_CL_paths.Append("Any")
        self.filter_CL_paths.AppendItems(CLs)
        self.filter_CL_paths.Select(0)
        
        self.filter_VehTypes_paths.SetItems([])
        self.filter_VehTypes_paths.Append("Any")
        self.filter_VehTypes_paths.AppendItems(CLs)
        self.filter_VehTypes_paths.Select(0)

        
        #self.list_VehTypes_filter_CLS.SetItems([])
        #self.list_VehTypes_filter_CLS.Append("Any")
        #self.list_VehTypes_filter_CLS.AppendItems(VehTypes)
        #self.list_VehTypes_filter_CLS.Select(0)
        
        self.__init_grid(self.grid_1)
        
        ColNames = ["PlateNo", "Type", "DetectionTime at CL#1 ", "DetectionTime at CL#2 /Code CL#1 ", "DetectionTime at CL#3"]
        self.grid_1.ClearGrid()
        self.grid_1.AppendCols(len(ColNames))
        [self.grid_1.SetColLabelValue(i, col) for i, col in enumerate(ColNames)]
        self.grid_1.AutoSizeColumns()
        
    
        #self.list_CL_filter_CLS.AppendItems(["One","Two","Three","Any"])
        #self.list_CL_filter_CLS.Select(0)
        #self.__init_grid(self.grid_1_CLS)
        #ColNames=["CLs set","Volume filtered"]
        #self.grid_1_CLS.ClearGrid()
        #self.grid_1_CLS.AppendCols(len(ColNames))
        #[self.grid_1_CLS.SetColLabelValue(i,col) for i,col in enumerate(ColNames)]
        
        #self.grid_1_CLS.AutoSizeColumns()

    def __file_dialog(self, wxstyle=wx.FD_SAVE):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        General function to support opening DB files, 
        optional param = wx.FileDialog style 
        wxstyle= wx.FD_SAVE | wx.FD_OPEN
        """
        
        
        dlg = wx.FileDialog(self,
                            message="Open",
                            defaultDir=os.getcwd(),
                            defaultFile="",
                            wildcard="Database files|*.db|Any file|*.*",
                            style=wxstyle)
        if dlg.ShowModal() == wx.ID_OK:
            filename = dlg.GetPath()
            if os.path.exists(filename):
                wx.MessageBox('Database already exists, overwrite?', 'Overwrite?', wx.YES_NO | wx.ICON_QUESTION | wx.STAY_ON_TOP)
        dlg.Destroy()
        
        self.__updateConsole('Database filename %(v1)s' % {'v1':filename})
        return filename
    
    def hh__sec(self, i):  

        try:
            res = int(i[0:2]) * 3600 + int(i[3:5]) * 60 + int(i[6:8])
        except:
            self.ErrMsg("Check time input format, correct one is: hh:mm:ss, i.e. 12:20:43")
            return 999999999999
        if len(i) > 8 or int(i[0:2]) > 24 or int(i[3:5]) > 60 or int(i[6:8]) > 60:
            self.ErrMsg("Check time input format, correct one is: hh:mm:ss, i.e. 12:20:43")
            return 999999999999
        else:
            return res
    
    def ___sec__hh(self, ss):
        try:
            [hh, mm] = divmod(ss, 3600)
        except:
            self.ErrMsg("Check time input format, correct one is: hh:mm:ss, i.e. 12:20:43")
            return
        if ss < 0:
            self.ErrMsg("Check time input format, correct one is: hh:mm:ss, i.e. 12:20:43")
            return
        [mm, ss] = divmod(mm, 60)
        res = []
        for i, el in enumerate([hh, mm, ss]):
            if el < 10:
                el = "0" + str(el)
            else:
                el = str(el)
            if i < 2:
                el = el + ":"
            res.append(el)
        res_str = res[0] + res[1] + res[2]
        return res_str, [hh, mm, ss]
    
    def __fill_grid(self, grid, filteresult, CLs=0): 
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        General function to fill grid with values from db.cur.fetchall() filtering result
        IN: wx.grid , db.cur.execute(str).fetchall() 
        """       
        try:
            filteresult
        except:
            self.ErrMsg("No filter selcted yet")
            return
        self.__init_grid(grid, True)
        
        for rowindex, row in enumerate(filteresult):
            grid.AppendRows()
            for colindex, col in enumerate(row):
                if colindex > 1 and CLs == 0:
                    col = self.___sec__hh(col)[0]
                if colindex==0 or colindex==1 or colindex==3 or colindex==4:
                    grid.SetReadOnly(rowindex,colindex,True)
                grid.SetCellValue(rowindex, colindex, str(col))
        
    def fill_grid_CLs(self): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        like __fill_grid() but fullfills self.grid_init with CLs data.
        run straight after db connection is established
        
        #TO DO SN: NADPISAC zeby aktualizowalo sie po process database 
        """ 
        self.DB.cur.execute("PRAGMA table_info(CountLocations)")
        self.__init_grid(self.grid_init)
        ColNames = [Col[1] for Col in self.DB.cur.fetchall()]
        self.grid_init.ClearGrid()
        self.grid_init.AppendCols(len(ColNames))
        [self.grid_init.SetColLabelValue(i, col) for i, col in enumerate(ColNames)]
        FilterResult = self.DB.cur.execute("select * from CountLocations")
        
        self.__fill_grid(self.grid_init, FilterResult, 1)
        
        FilterResult=self.DB.cur.execute('select FromClCode,ToCLCode,enabled,t0,tCur,mint,maxt from Matrix').fetchall()
        
        
        self.__fill_grid(self.grid_1_paths,FilterResult,1)
        
        numrows=self.grid_1_paths.GetNumberRows()
        choice_editor = wx.grid.GridCellChoiceEditor(['yes','no'], True) 
        for row in range(numrows):
            self.grid_1_paths.SetCellEditor(row, 2, choice_editor)
        
    def __init_grid(self, grid, rows=False):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####        
        clears the grid
        IN: grid=wx.grid, rows=boolean
        """
        if not rows:
            if grid.GetNumberCols() > 0:
                grid.DeleteCols(0, self.grid_1.GetNumberCols())        
        if grid.GetNumberRows() > 0:
            try:            
                grid.DeleteRows(0, grid.GetNumberRows())
            except:
                pass
            
    def __init_Matrix_Grid(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        sets wx.grid_1_mat row and col names + appends values to list_CL_filter_Mat      
        run after DB connection is established
        IN: DB table matrix
        """
        temptn = self.DB.cur.execute("PRAGMA table_info(Matrix)").fetchall()
        self.MatrixTablenames = [str(row[1]) for ind, row in enumerate(temptn)]
        
        numberOfCL = self.DB.cur.execute("select count (distinct CLCode) from CountLocations").fetchall()[0][0]
        tempCLCodes = self.DB.cur.execute("select CLCode from CountLocations").fetchall()
        CLCodes = [str(row[0]) for ind, row in enumerate(tempCLCodes)]
        
        self.__init_grid(self.grid_1_Mat)
        self.grid_1_Mat.ClearGrid()
        self.grid_1_Mat.AppendCols(len(CLCodes))
        self.grid_1_Mat.AppendRows(len(CLCodes))
        [self.grid_1_Mat.SetColLabelValue(i, col) for i, col in enumerate(CLCodes)]
        [self.grid_1_Mat.SetRowLabelValue(i, col) for i, col in enumerate(CLCodes)]
        
        choice = [self.MatrixTablenames[i] for i in range(3, len(self.MatrixTablenames))]
        self.list_CL_filter_Mat.AppendItems(choice)
        self.grid_1_Mat.AutoSizeColumns()
    def __Export_Excel(self, grid):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        general function to export GRID data to excel
        #TO DO SN: Sprobowac znalezc procedure ktoa od razu wklei cala tabele do excela: Range("A1:B1").value=x ?
        #nie zainstalowalem excela,
        w przykladach w necie jest podobnie do twojej metody, na przyklad:
        def addDataColumn(worksheet, columnIdx, data):
            range = worksheet.Range("%s:%s" % (
            genExcelName(0, columnIdx),
            genExcelName(len(data) - 1, columnIdx),
            ))
        for idx, cell in enumerate(range):
            cell.Value = data[idx]
    return range
        """
        
        self.__updateConsole('Exporting to Excel (col: %(v1)s x row: %(v2)s )' % {'v1':grid.GetNumberCols, 'v2':grid.GetNumberRows })
        nocol = grid.GetNumberCols()
        norow = grid.GetNumberRows()
        Content = [[grid.GetColLabelValue(i) for i in range(nocol)]]
        #TO DO RK: WKLEJ OD RAZU WSZYSTKO
        for row in range(norow):
            Content.append([grid.GetCellValue(row, col) for col in range(nocol)])
        try:
            import win32com.client
        except:
            import win32com.client  
        Excel = win32com.client.Dispatch("Excel.Application")   
        Excel.Visible = 1
        Excel.Workbooks.Add()    
        for i, col in enumerate(Content):
            Excel.Cells(i + 3, 1).Value = grid.GetRowLabelValue(i)
            for j, row in enumerate(col):               
                Excel.Cells(i + 2, j + 2).Value = str(row)

    def __handler_DB_init(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        DB inititialize constructor.
        Runs filedialog, established connection to new database,
        establishes self.DB as an instance of DataBase
        (see: DataBase.__init__)
        updates console
        populates filters
        fills grid with CL info
        initializes Matrix Grid
        #PROGRESSBAR=TRUE        
        
        """
        self.__updateConsole("Creating new database")
        
        filename = self.__file_dialog()

        try:
            os.remove(filename)
        except:
            pass

        # DONE SN: sprawdz, czy nie inicjalizujesz istniejacej - wiesza sie wtedy.
        self.DB = DataBase([self.Visum, filename,
                                      None,
                                      True,
                                      self.TSys_Combo.GetValue(),
                                      self.DSeg_Combo.GetValue(),
                                      self.Interpolate])
        self.__updateConsole("Database " + filename + " created")
        self.__updateConsole("Database tables created")
        self.__updateConsole("Data About CLs imported from Visum")
        self.__populate_filters()
                
        self.__updateConsole("Data imported")
        self.fill_grid_CLs()
        
        self.__init_Matrix_Grid()
        self.__updateConsole("CLs skim matrix initialized")
        self.__updateConsole("DataBase initialization complete, you may play with the data")        
       
    def __handler_DB_connect(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        DB connection constructor.
        Runs filedialog, established connection to existing database,
        establishes self.DB as an instance of DataBase
        (see: DataBase.__init__)
        updates console
        populates filters
        fills grid with CL info
        initializes Matrix Grid
        #PROGRESSBAR=TRUE
        """
        try: 
            self.DB.con
            caption = "Connection with database is active, restart the script to connect to another database"
            dlg = wx.MessageDialog(self, caption, "i2 APNR", wx.OK)
            result = dlg.ShowModal() == wx.ID_YES
            dlg.Destroy()
        except:
            
            self.__updateConsole("Connecting to existing Database")
            filename = self.__file_dialog(wxstyle=wx.FD_OPEN)
            self.__updateConsole("Database " + filename + " connection established")
            
            #filename='C:\9i9i9i.db'
            self.__updateConsole("Database tables created")
            self.__updateConsole("Data About CLs imported from Visum")
            
            self.DB = DataBase([self.Visum,
                                          filename,
                                          None,
                                          False,
                                          self.TSys_Combo.GetValue(),
                                          self.DSeg_Combo.GetValue(),
                                          self.Interpolate])
            #print self.DB.TSys
            self.__populate_filters()
            
            self.__updateConsole("Data imported")
            self.fill_grid_CLs()
            self.__init_Matrix_Grid()
            self.__updateConsole("CLs skim matrix initialized")
            self.__updateConsole("DataBase initialization complete, you may play with the data")
        
    def __handler_import(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Data importer.
        Runs dir dialog, runs Txt_to_DB
        establishes self.DB as an instance of DataBase
        (see: DataBase.Txt_to_DB)
        #PROGRESSBAR=TRUE
        
        #TO DO SN: Dwa rodzaje inputow (txt,xls) wybor zmiennej w kodzie, nie przez uzytkownika - bo kazdy importer sprzedajemy osobno. 
        #jezeli drugi importer bedzie w pelni sprawny, to jest pietnascie sekund pracy
        Uruchom procedure Txt_To_DB z odpowiednim parametrem.
        """
        
        self.__updateConsole('Running data importer')
        
        caption = "This operation will add results from text files to database,\ndo you want to continue?"
        dlg = wx.MessageDialog(self, caption, "i2 APNR", wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal() == wx.ID_YES
        dlg.Destroy()
        if not result:
            return  
        
        
        dlg = wx.DirDialog(self, "Choose a directory:", defaultPath=os.getcwd(), style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
        dlg.Destroy()
        try:
            self.DB
        except:
            self.ErrMsg("Database connection not established yet, \nplease choose database to connect with, or create new one.")
            return
        #print path
        
        if self.Importer=="ARGUS":
            self.DB.Txt_to_DB2(path) #nie ma tego starego importera
        elif self.Importer=="PK":
            self.DB.XLS_to_DB(path)
        self.__populate_filters()
        
    def __handler_filter(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Main handler to pass queries to DB
        Takes data from respective filters (CountLocations, VehType, FromTime,ToTime,PlateNo)
        and passes them to DB.Filter
        Calculates characteristics of selected query self.Calc_Charactersistics       
        """
        '''CLs'''
        
        
        selections = [self.list_CL_filter.GetSelection(), self.list_CL_filter2.GetSelection(), self.list_CL_filter3.GetSelection()]
        filter_CLs = [str(self.list_CL_filter.GetString(selection)) for selection in selections]
        
        if filter_CLs[0] == 'Any' and (filter_CLs[1] != 'Any' or filter_CLs[2] != 'Any'):
            self.ErrMsg("Incorrect selection")
        
        if filter_CLs[0] != 'Any' and filter_CLs[1] == 'Any' and filter_CLs[2] != 'Any':
            self.ErrMsg("Incorrect selection")
            
        if filter_CLs[0] != 'Any' and filter_CLs[1] != 'Any' and filter_CLs[2] == 'Any':
            filter_CLs = [filter_CLs[0], filter_CLs[1]]
            
        if filter_CLs[0] != 'Any' and filter_CLs[1] == 'Any':
            filter_CLs = [filter_CLs[0]]
            
        if filter_CLs[0] == 'Any' and filter_CLs[1] == 'Any' and filter_CLs[2] == 'Any':
            filter_CLs = ['None']
            
            
        
        if 'None' in filter_CLs and len(filter_CLs) > 1:
            self.ErrMsg("None as a value for CL filter can be selected only as a single selection.")
            return
        if filter_CLs == ['None']:
            filter_CLs = None
            
        
            
        '''VehTypes'''
        selections = [a for a in self.list_VehTypes_filter.GetSelections()]
        filter_VehTypes = [str(self.list_VehTypes_filter.GetString(selection)) for selection in selections]
        if 'None' in filter_VehTypes and len(filter_VehTypes) > 1:
            self.ErrMsg("None as a value for VehType filter can be selected only as a single selection.")
            return
        if len(filter_VehTypes) == 1: filter_VehTypes = filter_VehTypes[0]
        
        if filter_VehTypes == 'Any': filter_VehTypes = None
        
        '''PlateNo'''
        filter_PlateNo = str(self.plate_no_text.GetValue())
        if filter_PlateNo == 'None':
            filter_PlateNo = None
        '''FromTime ToTime'''
        filter_FromTime = self.from_time_text.GetValue()
        filter_ToTime = self.to_time_text.GetValue()        
        filter_FromTime = self.hh__sec(filter_FromTime)
        filter_ToTime = self.hh__sec(filter_ToTime)
        if filter_FromTime == 999999999999:
            return
        if filter_ToTime == 999999999999:
            return
        
        
        
        
        try:
            self.__updateConsole('Data filtered CLs: %(v1)s VehType: %(v2)s PlateNo: %(v3)s From/To Time: %(v4)s/%(v5)s' % {'v1': str(len(filter_CLs)), 'v2': str(filter_VehTypes), 'v3':str(filter_PlateNo), 'v4': str(filter_FromTime), 'v5': str(filter_ToTime)})
        except:
            pass
        self.Filter_Result = self.DB.Filter(False, filter_CLs, filter_VehTypes, filter_PlateNo, filter_FromTime, filter_ToTime)
        self.__fill_grid(self.grid_1, self.Filter_Result)
        if filter_PlateNo == None:        
            self.Calc_Characteristics(True)
        c = self.Characteristics
        self.grid_stats.SetCellValue(0, 0, str(c[0]))
        self.grid_stats.SetCellValue(0, 1, str(c[5]))
        self.grid_stats.SetCellValue(0, 2, str(c[1]))
        self.grid_stats.SetCellValue(0, 3, str(c[2]))
        self.grid_stats.SetCellValue(0, 4, str(c[3]))
        self.grid_stats.SetCellValue(0, 5, str(c[4]))

        
    #ponizej nie potrzebne?
    
    def __handler_GUI_Plot(self, event):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Plotter.
        Gets selection from self.combo_box_1
        Initializes PlotPanel
        (see: PlotPanel.__init__)
        adds subplot
        plots respective figure
        creates self.Points - used in Excel Export
        
        
        #TO DO SN+RK: inicjalizacja nowego wykresu tak,zeby nie migal przy odswiezeniu. 
        Trzeba wczesniej strorzyc Figure i dodawac tylko plot i clear?
        aktualnie nierozwiazywalny klopot, nie znalazlem infa na ten temat
        
        """
        
        self.__updateConsole('plotting data from selection: ' + self.combo_box_1.GetValue())
        
        try:
            self.Filter_Result
        except:
            self.ErrMsg("No data filtered from DB")
        sel = self.combo_box_1.GetSelection()
        self.PlotPanel = PlotPanel(self.plot_support_panel)
        self.Figure = self.PlotPanel.figure
        self.Subplot = self.Figure.add_subplot(111)
        
        if sel == 0:
            #0# Simple for single CL
            ys = [x[2] for x in self.Filter_Result]
            self.Points = [range(len(ys)), ys]
            self.Subplot.plot(self.Points[0], self.Points[1])
            
        elif sel == 1:
            #1# Histogram single CL
            ys = [x[2] for x in self.Filter_Result]
            self.Points = [range(len(ys)), ys]
            self.Subplot.hist(self.Points[1], bins=40)
            #self.Subplot.refresh()
        elif sel == 2:
            #2# Double CL travel time (t)
            ys = [x[3] - x[2] for x in self.Filter_Result]
            self.Points = [range(len(ys)), ys]
            self.Subplot.plot(self.Points[0], self.Points[1])
        elif sel == 3:
            #3# Histogram double CL
            ys = [x[3] - x[2] for x in self.Filter_Result]
            self.Points = [range(len(ys)), ys]
            self.Subplot.hist(self.Points[1], bins=40)
        elif sel == 4:
            #4# Frequencies
            ys = [self.Filter_Result[i + 1][2] - self.Filter_Result[i][2] for i, a in enumerate(self.Filter_Result[:-1])]
            self.Points = [range(len(ys)), ys]
            self.Subplot.plot(self.Points[0], self.Points[1])
        
    def __handler_excel_plot_export(self, event): 
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Exports plot [x,y] to Excel
        """
        
        self.__updateConsole('Exporting plot data to Excel')
        try:
            self.Points
        except:
            self.ErrMsg("No Plot data to export. Generate plot first.")
            return
        
        import win32com.client   
        Excel = win32com.client.Dispatch("Excel.Application")   
        Excel.Visible = 1
        Excel.Workbooks.Add()
        Excel.Cells(1, 2).Value = "x"
        Excel.Cells(1, 3).Value = "y"
        for i, col in enumerate(self.Points[0]):
            Excel.Cells(i + 2, 2).Value = str(col) 
        for i, col in enumerate(self.Points[1]):
            Excel.Cells(i + 2, 3).Value = str(col)
            
        

    def __handler_CLs_click(self, event):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Clears Visum Marking
        Marks respective CL (Clicked Row from self.grid_init) in Visum
        """
        
        self.__updateConsole('marking respective CL in Visum')
        self.grid_init.SelectRow(event.GetRow())
        self.Visum.Net.Marking.Clear()        
        self.Visum.Net.Marking.ObjectType = 11
        key = int(self.grid_init.GetCellValue(event.GetRow(), 0))
        try:
            self.Visum.Net.Marking.Add(self.Visum.Net.CountLocations.ItemByKey(key))
        except:
            pass
    
    
    def __handler_savePthtoDb(self,event):
        
        rows=self.grid_1_paths.GetNumberRows()
        for row_ind in range(rows):
            fullrow=[]
            for col_ind in range(7):
                cell=self.grid_1_paths.GetCellValue(row_ind,col_ind)
                fullrow.append(cell)
            self.DB.cur.execute('update Matrix set enabled = ?, mint=?, maxt = ? where FromCLCode=? and ToCLCode= ?',(fullrow[2],fullrow[5],fullrow[6],fullrow[0],fullrow[1]))
        
        self.DB.con.commit()
    
    def handler_filtrujPth(self,event):
        selections = [self.filter_CL_paths.GetSelection(), self.filter_VehTypes_paths.GetSelection()]
        filter_CLs = [str(self.list_CL_filter.GetString(selection)) for selection in selections]
        
        #print filter_CLs
        
        if filter_CLs[0] != 'Any' and filter_CLs[1] != 'Any':
            filter_CLs = [filter_CLs[0], filter_CLs[1]]
            res=self.DB.cur.execute('SELECT FromCLCode,ToCLCode,enabled,T0,TCur,mint,maxt FROM matrix WHERE FromCLCode = ? and ToCLCode = ?',filter_CLs).fetchall()
            self.__fill_grid(self.grid_1_paths, res,1)
        
        if filter_CLs[0] == 'Any' and filter_CLs[1] != 'Any':
            filter_CLs = [filter_CLs[0], filter_CLs[1]]
            res=self.DB.cur.execute('SELECT FromCLCode,ToCLCode,enabled,T0,TCur,mint,maxt FROM matrix WHERE FromCLCode <> ? and ToCLCode = ?',filter_CLs).fetchall()
            self.__fill_grid(self.grid_1_paths, res,1)
            
        if filter_CLs[0] != 'Any' and filter_CLs[1] == 'Any':
            filter_CLs = [filter_CLs[0], filter_CLs[1]]
            res=self.DB.cur.execute('SELECT FromCLCode,ToCLCode,enabled,T0,TCur,mint,maxt FROM matrix WHERE FromCLCode = ? and ToCLCode <> ?',filter_CLs).fetchall()
            self.__fill_grid(self.grid_1_paths, res,1)
            
        if filter_CLs[0] == 'Any' and filter_CLs[1] == 'Any':
            filter_CLs = [filter_CLs[0], filter_CLs[1]]
            res=self.DB.cur.execute('SELECT FromCLCode,ToCLCode,enabled,T0,TCur,mint,maxt FROM matrix').fetchall()
            self.__fill_grid(self.grid_1_paths, res,1)
        
    def __handler_Path_click(self, event):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Clears Visum Marking
        Marks respective pair of CLs (Clicked Row and Col from self.grid_1_Mat) in Visum
        Tries to mark respective Path between O and D        
        """
        self.grid_1_paths.SelectRow(event.GetRow())
        self.Visum.Net.Marking.Clear()
        self.Visum.Net.Marking.ObjectType = 11
        key1 = int(self.grid_1_paths.GetCellValue(event.GetRow(), 0))        
        self.Visum.Net.Marking.Add(self.Visum.Net.CountLocations.ItemByKey(key1)) 
        key2 = int(self.grid_1_paths.GetCellValue(event.GetRow(), 1))        
        self.Visum.Net.Marking.Add(self.Visum.Net.CountLocations.ItemByKey(key2))      
        try:
            path=self.Visum.Net.Paths.ItemByKey(12,int(100000*key1+key2))
            
            self.Visum.Net.Marking.ObjectType = 19
            self.Visum.Net.Marking.Add(path)
        except:
            pass
                             
        
    
    def __handler_Mtx_click(self, event):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Clears Visum Marking
        Marks respective pair of CLs (Clicked Row and Col from self.grid_1_Mat) in Visum        
        """
        self.Visum.Net.Marking.Clear()
        self.Visum.Net.Marking.ObjectType = 11        
        self.grid_1_Mat.SelectBlock(event.GetRow(), event.GetCol(), event.GetRow(),event.GetCol())        
        key = int(self.grid_init.GetCellValue(event.GetRow(), 0))
        self.Visum.Net.Marking.Add(self.Visum.Net.CountLocations.ItemByKey(key)) 
        key = int(self.grid_init.GetCellValue(event.GetCol(), 0))
        self.Visum.Net.Marking.Add(self.Visum.Net.CountLocations.ItemByKey(key))
        return
     
    def __handler_filter_ST(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Main handler to multi filter data basa.
        IN: selection in self.list_CL_filter_CLS - single, double, triple(not yet!)
        IN: VehTypes: self.list_VehTypes_filter_CLS
        OUT: fills self.grid_1_CLS with len(fetchall())
        #PROGRESBAR=TRUE
        #DONE SN: ZLY FORMAT WYNIKU DZIALANIA FILTRA DLA POJEDYNCZEGO CL i dla zadnego CL (PlateNo)
        """
        '''
        selections=[a for a in self.list_VehTypes_filter_CLS.GetSelections()]
        filter_VehTypes=[str(self.list_VehTypes_filter_CLS.GetString(selection)) for selection in selections]
        selections=[a for a in self.list_CL_filter_CLS.GetSelections()]
        if 'Any' in filter_VehTypes and len(filter_VehTypes)>1:
            self.ErrMsg("None as a value for VehType filter can be selected only as a single selection.")
            return
        if len(filter_VehTypes)==1: filter_VehTypes=filter_VehTypes[0]
        
        if filter_VehTypes=='Any': filter_VehTypes=None
        filter_FromTime=int(self.from_time_text_CLS.GetValue())
        filter_ToTime=int(self.to_time_text_CLS.GetValue())
        
        CLs=self.DB.cur.execute("select CLCode from CountLocations").fetchall()
        CLs=[str(CL[0]) for CL in CLs]
        """
        Statystyki dla pojedynczych CLs, liczy Volume, VOL_Errors
        """
        
        self.dialog = wx.ProgressDialog ( 'Progress', "Gathering data from DB - CL pairs", maximum = len(CLs)+1 )
        u=1
        CLsquare=len(CLs)**2
        for CL1 in CLs:
            for CL2 in CLs:
                if CL1!=CL2:
                        self.Filter_Result=self.DB.Filter(False,[CL1,CL2],filter_VehTypes,None,filter_FromTime,filter_ToTime)
                        if len(self.Filter_Result)>0:
                           [c_len,c_min,c_mean,c_mod,c_max,a]=self.Calc_Characteristics()
                           self.dialog.Update(u)
                           Query="UPDATE Matrix SET APNR_VOLUME_ANY = '"+str(c_len)+"', APNR_TMIN_ANY='"+str(c_min)+"', APNR_TMEAN_ANY='"+str(c_mean)+"', APNR_TMOD_ANY='"+str(c_mod)+"', APNR_TMAX_ANY='"+str(c_max)+"' where FromCLCode= '"+CL1+"' and ToCLCode= '"+CL2+"'"
                           self.DB.con.execute(Query)
                           paste=["["+CL1+" , "+CL2+"]",str(c_len)]
                           self.grid_1_CLS.AppendRows()
                           self.grid_1_CLS.SetCellValue(self.grid_1_CLS.GetNumberRows()-1,0,str(paste[0]))
                           self.grid_1_CLS.SetCellValue(self.grid_1_CLS.GetNumberRows()-1,1,str(paste[1]))
                        u+=1 
                                          
        self.DB.con.commit() 
        self.dialog.Destroy() 
        '''
        
    def __handler_Export_Filter(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Passes grid to self.__Export_Excel
        """
        self.__Export_Excel(self.grid_1)
        self.__updateConsole('Exporting filter results to Excel')
        
    def __handler_Export_CL(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Passes grid to self.__Export_Excel
        """
              
        self.__Export_Excel(self.grid_init)
        self.__updateConsole('Exporting CLs stats to Excel')

    def __handler_export_matrix(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Passes grid to self.__Export_Excel
        """
        self.__Export_Excel(self.grid_1_Mat)
        
    def __handler_export_Statistics(self, event):
        self.__Export_Excel(self.grid_stats)
        
        self.__updateConsole('Exporting statistics to Excel')
        
    def __handler_Export_Paths_2_Visum(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Creates paths in Visum from table Matrix APNR_VOLUME_OD
        
        IN: list_CL_filter_Mat
        OUT: Visum paths
        
        SEE: Database.Make_Paths(type)
        
        #DONE RK PROGRESSBAR
        """
        self.DB.set_Paths(self.Paths)
        self.__updateConsole("poczatek __handler_Export_Paths_2_Visum")  
        sel = self.list_CL_filter_Mat.GetSelection()
        self.__updateConsole("sel:"+str(sel)) 
        dic = {2:0, 3:1, 4:2, 5:3}
        
        try:            
            self.DB.Make_Paths(dic[sel])
        except:
            self.DB.Make_Paths()
        
        self.__updateConsole('Creating paths in Visum from table Matrix APNR_VOLUME_OD')
    
            
    def __handler_fill_matrix(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Main handler to multi filter data basa.
        IN: selection in self.list_CL_filter_CLS - single, double, triple(not yet!)
        IN: VehTypes: self.list_VehTypes_filter_CLS
        IN: DB table Matrix
        
        OUT: self.grid_1_mat -> fill
        """
        selection = self.list_CL_filter_Mat.GetSelection()
        SelectedStr = str(self.list_CL_filter_Mat.GetString(selection))
        tempCLCodes = self.DB.cur.execute("select CLCode from CountLocations").fetchall()
        CLCodes = [str(row[0]) for ind, row in enumerate(tempCLCodes)]
        ClCod = len(CLCodes)
        FilterResult = self.DB.cur.execute("select " + SelectedStr + " from Matrix")
        A = FilterResult.fetchall() 
        S = [A[i * ClCod:(i + 1) * ClCod] for i in range(ClCod)]
        for rowindex, row in enumerate(CLCodes):
            for colindex, col in enumerate(CLCodes):
                self.grid_1_Mat.SetCellValue(rowindex, colindex, str(S[rowindex][colindex][0]))
        self.Colour_Mtx()
        
        self.__updateConsole('Filling grid by data from Matrix')

    def __handler_calc_matrix(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Calculates Values in DB table Matrix
        
        IN: list_CL_filter_Mat
        OUT: INSERT INTO MATRIX
        
        SEE: Database.Populate_Matrix_from_Visum(type)
        #DONE SN: Jesli zmienisz cokolwiek w definicji tabeli Matrix, sprawdz indexy w tej procedurze!!!!
        
        """
        
        sel = self.list_CL_filter_Mat.GetSelection()
              
        dic = {2:0, 3:1, 4:2, 5:3}
        if sel <= 5:
           """VISUM MATRICES"""
           if sel == 0:
               """STATE"""
               self.DB.Populate_Matrix_from_Visum(0, 1)
           elif sel == 1:
               """VOLUME_VISUM"""                           
               self.DB.Populate_Matrix_from_Visum(4)
               return
           else:
                """TO,TCUR,IMP,DIST"""
                self.DB.Populate_Matrix_from_Visum(sel - 2)
        if sel in [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]:
            """VOLUME_APNR_0, TMIN, TMAX, TMEAN, TMOD"""
            self.ErrMsg("Those matrices are automatically calculated during Processing Database")
            return
            
        if sel == 17:
            """PATHNODES"""            
            self.DB.Populate_Matrix_from_Visum(-1, 1)
        if sel in [21, 22]:
            """IS CONTAINED IN CONTAINS"""            
            self.DB.Licz_Zaleznosci_Miedzy_Rejonami()
               
            
        

    def __handler_help(self, event): # wxGlade: APNR_GUI.<event_handler>
        self.__updateConsole('Help clicked')
        os.startfile(self.Paths["Help"])
    def __handler_cancel_click(self, event): # wxGlade: APNR_GUI.<event_handler>
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Destroys self
        """
        self.Destroy()

    def __handler_Export_Visum_Zones(self, event):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Creates zones in Visum DB Table CountLocations
        
        IN: CountLocations: CLCode, WKT,FromNodeNo,ToNodeNo 
        OUT: Visum zones, 
            1 zone per CL, 
            two connectors (one origin connector - Zone -> FromNodeNo)
            one dest connector - FromNodeNo-> Zone
        
        
        #DONE RK+SN: ErrMsg - Tak/Nie
        """
        
        
        caption = "This operation will delete all zones in your Version file,\ndo you want to continute"
        dlg = wx.MessageDialog(self, caption, "i2 APNR", wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal() == wx.ID_YES
        dlg.Destroy()
        if not result:
            return       
        ZoneNos = self.Visum.Net.Zones.GetMultiAttValues("No")
        CLs = self.DB.cur.execute("select * from CountLocations").fetchall()
        self.dialog = wx.ProgressDialog ('Progress', "Creating Visum Zones", maximum=len(CLs) + 1)
        print ZoneNos
        for ZoneNo in ZoneNos:
            self.Visum.Net.RemoveZone(self.Visum.Net.Zones.ItemByKey(ZoneNo[1]))
        self.dialog.Update(1)
        
        CLs = self.DB.cur.execute("select * from CountLocations").fetchall()
        u = 1
        for CL in CLs:
            u += 1
            self.dialog.Update(u)
            A = CL[2]
            X = float(A[6:A.index(" ") - 1])
            Y = float(A[A.index(" ") + 1:-1])
            self.Visum.Net.AddZone(CL[0], X, Y)
            self.Visum.Net.Zones.ItemByKey(CL[0]).SetAttValue("Name", CL[1])
            #DONE SN: JESLI bylyby problemy z create zones, jakies nielogiczne rozkladyby zglosil, czy cos, to sprawdz ponizej
            self.Visum.Net.AddConnector(CL[0], CL[3]) #stworz konektor
            self.Visum.Net.AddConnector(CL[0], CL[4])
            self.Visum.Net.Connectors.SourceItemByKey(CL[0], CL[4]).SetAttValue("TSysSet", None) #zablokuj konektor w jednym kierunku
            self.Visum.Net.Connectors.DestItemByKey(CL[3], CL[0]).SetAttValue("TSysSet", None)
        self.dialog.Destroy()
        
        
    def __handler_export_Visum_Matrix(self, event):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Export VOLUMES_APNR to Visum matrix no: 1212
        
        IN: DB Matrix Table: APNR_VOLUME_OD 
        OUT: Visum.Net.Matrices.ItemByKey(1212)
        
        #DONE RK: sprawdzic wartosci
        
        """
        try:
            self.Visum.Net.AddODMatrix(1212)
        except:
            pass
        COL = self.list_CL_filter_Mat.GetStringSelection()
        tempCLCodes = self.DB.cur.execute("select CLCode from CountLocations").fetchall()
        CLCodes = [str(row[0]) for ind, row in enumerate(tempCLCodes)]
        ClCod = len(CLCodes)
        Query = "select " + COL + " from Matrix"
        FilterResult = self.DB.cur.execute(Query)
        A = FilterResult.fetchall()
        B = []
        test = False
        for a in A:
            try:
                B.append(int(a[0]))
            except:
                test = True
                B.append(0)
        if test:         
            self.ErrMsg("Non numeric Values in Matrix, check again.\nPS. Value='None' is acceptable.\n Check data.")
        S = [B[i * ClCod:(i + 1) * ClCod] for i in range(ClCod)]
        Mtx = self.Visum.Net.Matrices.ItemByKey(1212)
        Mtx.SetValues(S)
    
    def __handler_Import_Skim_Min(self,event):
        self.MinMax="mint"
        DLG = ImportMtxDialog(self)
        DLG.Show() 
        
    def __handler_Import_Skim_Max(self,event):
        self.MinMax="maxt"
        DLG = ImportMtxDialog(self)
        DLG.Show() 
    
    def __handler_process_menu(self, event):
        DLG = Process_Dialog(self)
        DLG.Show()
        
    def __handler_fratar(self, event):        
       try:
           self.DB
       except:
            self.ErrMsg("No DB connection established") 
            return
       self.__get_Fratar_Vols()
       self.DB.Fratar()    
        
    def __get_Fratar_Vols(self):
        try:
            A=self.Visum.Net.CountLocations.GetMultipleAttributes(["Name","i2_APNR_VOL_FRATAR_FROM"])
            B=self.Visum.Net.CountLocations.GetMultipleAttributes(["Name","i2_APNR_VOL_FRATAR_TO"])
            for a in A:
                self.DB.con.execute("UPDATE COUNTLOCATIONS SET VOL_FRATAR_FROM = "+ str(a[1]) +" where CLCODE= "+ str(a[0]))
            for b in B:
                self.DB.con.execute("UPDATE COUNTLOCATIONS SET VOL_FRATAR_TO = "+ str(b[1]) +" where CLCODE= "+ str(b[0]))
            self.DB.con.commit()            
                
            
        except:
            self.ErrMsg("Please define \"i2_APNR_VOL_FRATAR_From\" and \"i2_APNR_VOL_FRATAR_To\" UDA for CLs first")
    
    def Calc_Characteristics(self, full=False):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        General calculation of basic characteristics of DB Query
        IF FULL==FALSE: tylko najpotrzebniejsze,
        IF FULL: wszystko

        """
        
        self.__updateConsole('General calculation of basic charecteristics of DB Query')
        unknownveh = "none"
        C_Filter_Result = self.Filter_Result
        if len(C_Filter_Result) < 3:
            self.Characteristics = ["-", "-", "-", "-", "-", "-"]
            return self.Characteristics
        if len(C_Filter_Result[0]) == 4:            
            ys = [C_Filter_Result[i][3] - C_Filter_Result[i][2] for i, a in enumerate(C_Filter_Result)]
        if len(C_Filter_Result[0]) == 5: 
            ys = [C_Filter_Result[i][4] - C_Filter_Result[i][2] for i, a in enumerate(C_Filter_Result)]   
        else:
            ys = [C_Filter_Result[i][2] for i, a in enumerate(C_Filter_Result)]
            pl = [C_Filter_Result[i][0] for i, a in enumerate(C_Filter_Result)]
            unknownveh = pl.count('-')
            
#        m=mode([1,2,3,4])
#        if m[1][0]>1:
#            m=m[0][0]
#        else:
#            m="None"
        if full:
            self.Characteristics = [len(ys),
                                  min(ys),
                                  numpy.mean(ys),
                                  - 1,
                                  max(ys),
                                  unknownveh,
                                  max(ys) - min(ys),
                                  numpy.std(ys),
                                  numpy.median(ys),
                                  0, #numpy.percentile(ys, 80),
                                  0, #numpy.percentile(ys, 95),
                                  0]#skew(ys)]    
        else:
            self.Characteristics = [len(ys),
                                  min(ys),
                                  numpy.mean(ys),
                                  - 1,
                                  max(ys), unknownveh]
        return self.Characteristics
                              
    def Colour_Mtx(self):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Colours self.grid_1_Mat.
        Each state has its own colour.
        """

        dic = {"null path":"#FFFFFF",
               "None":"#FFFFFF",
                "ok":"#A3A948",
               "tail loop":"#EDB92E",
               "both loop":"#EDB92E",
               "head loop":"#EDB92E",
               "diag":"#F85931",
               "reverse":"#009989",
               "no SP found":"#CE1836"}
        tempCLCodes = self.DB.cur.execute("select CLCode from CountLocations").fetchall()
        CLCodes = [str(row[0]) for ind, row in enumerate(tempCLCodes)]
        ClCod = len(CLCodes)
        FilterResult = self.DB.cur.execute("select STATE from Matrix")
        A = FilterResult.fetchall() 
        S = [A[i * ClCod:(i + 1) * ClCod] for i in range(ClCod)]
        for rowindex, row in enumerate(CLCodes):
            for colindex, col in enumerate(CLCodes):
                self.grid_1_Mat.SetCellBackgroundColour(rowindex, colindex, dic[str(S[rowindex][colindex][0])])
        
    def ErrMsg(self, message):
        """
        ###
        Automatic Plate Number Recognition Support
        (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
        ####
        Main function to create error message box
        """
        wx.MessageBox(message, "i2 APNR Error", style=wx.OK | wx.ICON_ERROR)

        

 
# end of class APNR_GUI

    def Update_DSeg(self, event):
        try:
            self.DB.DSeg = self.DSeg_Combo.GetValue()
            
        except:
            pass
    def Update_TSys(self, event):
        try:
            self.DB.TSys = self.TSys_Combo.GetValue()
            
        except:
            pass
    def Update_Interpolate(self, event):
        
        if not self.Interpolate:
            self.Interpolate = True
        else:
            self.Interpolate = False  
        try:
            self.DB.Interpolate = self.Interpolate
        except:
            pass
        #self.DB.Licz_Zaleznosci_Miedzy_Rejonami()
        #self.DB.Licz_Nowe_Volumes()
        
        
    def Make_File_Paths(self):        
        self.Paths = {}
        try:
            self.Paths["MainVisum"] = self.Visum.GetWorkingFolder()
        except:
            return
        self.Paths["ScriptFolder"] = self.Paths["MainVisum"] + "\\AddIns\\intelligent-infrastructure\\APNR"
        self.Paths["Logo"] = self.Paths["ScriptFolder"] + "\\help\\i2_logo.png"
        self.Paths["Help"] = self.Paths["ScriptFolder"] + "\\Help\\help.htm" 
        self.Paths["Report"] = self.Paths["ScriptFolder"] + "\\report.txt"
        self.Paths["Exclusions"] = self.Paths["ScriptFolder"] + "\\exclusions.txt"

class Process_Dialog(wx.Dialog):
    def __init__(self, parent, *args, **kwds):    
        # begin wxGlade: ImportMtxDialog.__init__
        kwds["style"] = wx.DEFAULT_DIALOG_STYLE
        wx.Dialog.__init__(self, parent, *args, **kwds)
        self.logo_copy = wx.StaticBitmap(self, -1, wx.Bitmap(self.Parent.Paths["Logo"], wx.BITMAP_TYPE_ANY))
        self.from_time_text_dlg = wx.TextCtrl(self, -1, "06:00:00")
        self.from_time_sizer_copy_staticbox = wx.StaticBox(self, -1, "from time")
        self.to_time_text_dlg = wx.TextCtrl(self, -1, "07:00:00")
        self.to_time_sizer_copy_staticbox = wx.StaticBox(self, -1, "to time")
        self.list_VehTypes_filter_dlg = wx.ListBox(self, -1, choices=[], style=wx.LB_MULTIPLE | wx.LB_NEEDED_SB)
        self.filter_VehType_sizer_copy_staticbox = wx.StaticBox(self, -1, "vehicle types")
        self.label_1_copy_1 = wx.StaticText(self, -1, "Process trips:")
        self.label_1 = wx.StaticText(self, -1, "1. exclude trips with travel time below ith and above jth percentile")
        self.ith_perc = wx.TextCtrl(self, -1, "0")
        self.from_time_sizer_copy_copy_staticbox = wx.StaticBox(self, -1, "i")
        self.to_time_text_dlg_copy = wx.TextCtrl(self, -1, "100")
        self.jth_perc_staticbox = wx.StaticBox(self, -1, "j")
        self.label_1_copy = wx.StaticText(self, -1, "2. cut trips at \"excluded\" CL pairs?")
        self.exclusion_check_box = wx.CheckBox(self, -1, "")
        self.label_1_copy_copy = wx.StaticText(self, -1, "3. cut at duplicate CLs")
        self.duplicate_check_box = wx.CheckBox(self, -1, "")
        self.label_1_copy_copy_copy = wx.StaticText(self, -1, "4. cut at stopovers longer than: x minutes.")
        self.stopover_textbox = wx.TextCtrl(self, -1, "1000")
        self.sizer_4_staticbox = wx.StaticBox(self, -1, "Define time interval and vehicle type")
        self.CancelBtn_dlg = wx.Button(self, -1, "Cancel")
        self.DLG_Process_Button = wx.Button(self, -1, "Process")

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.Cancel_Click, self.CancelBtn_dlg)
        self.Bind(wx.EVT_BUTTON, self.Calc_Click, self.DLG_Process_Button)
        # end wxGlade

        try:
            self.Parent.DB
        except:
            self.Parent.ErrMsg("Connect to DB first")
            self.Destroy()
            return
        self.Parent.DB.cur.execute("Select Distinct VehType from DetectedVehicles")
        VehTypes = self.Parent.DB.cur.fetchall()
        VehTypes = [VehType[0] for VehType in VehTypes]
        self.list_VehTypes_filter_dlg.Append("None")
        self.list_VehTypes_filter_dlg.AppendItems(VehTypes)
        self.list_VehTypes_filter_dlg.Select(0)
        

    def __set_properties(self):
        # begin wxGlade: ImportMtxDialog.__set_properties
        self.SetTitle("Process DB")
        self.SetSize((447, 575))
        self.logo_copy.SetMinSize((-1, 20))
        self.logo_copy.SetBackgroundColour(wx.Colour(240, 240, 240))
        self.label_1_copy_1.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.BOLD, 0, ""))
        self.exclusion_check_box.SetValue(1)
        self.duplicate_check_box.SetValue(1)
        self.stopover_textbox.SetMinSize((10, 20))
        self.CancelBtn_dlg.SetMinSize((87, -1))
        self.DLG_Process_Button.SetMinSize((87, -1))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: ImportMtxDialog.__do_layout
        sizer_3 = wx.BoxSizer(wx.VERTICAL)
        sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
        self.sizer_4_staticbox.Lower()
        sizer_4 = wx.StaticBoxSizer(self.sizer_4_staticbox, wx.VERTICAL)
        sizer_10 = wx.BoxSizer(wx.VERTICAL)
        sizer_11_copy_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_11_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_11 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_7_copy = wx.BoxSizer(wx.HORIZONTAL)
        self.jth_perc_staticbox.Lower()
        jth_perc = wx.StaticBoxSizer(self.jth_perc_staticbox, wx.HORIZONTAL)
        self.from_time_sizer_copy_copy_staticbox.Lower()
        from_time_sizer_copy_copy = wx.StaticBoxSizer(self.from_time_sizer_copy_copy_staticbox, wx.HORIZONTAL)
        sizer_8 = wx.BoxSizer(wx.HORIZONTAL)
        self.filter_VehType_sizer_copy_staticbox.Lower()
        filter_VehType_sizer_copy = wx.StaticBoxSizer(self.filter_VehType_sizer_copy_staticbox, wx.HORIZONTAL)
        sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
        self.to_time_sizer_copy_staticbox.Lower()
        to_time_sizer_copy = wx.StaticBoxSizer(self.to_time_sizer_copy_staticbox, wx.HORIZONTAL)
        self.from_time_sizer_copy_staticbox.Lower()
        from_time_sizer_copy = wx.StaticBoxSizer(self.from_time_sizer_copy_staticbox, wx.HORIZONTAL)
        sizer_3.Add(self.logo_copy, 1, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        from_time_sizer_copy.Add(self.from_time_text_dlg, 1, wx.ALL, 5)
        sizer_7.Add(from_time_sizer_copy, 1, wx.EXPAND, 0)
        to_time_sizer_copy.Add(self.to_time_text_dlg, 1, wx.ALL, 5)
        sizer_7.Add(to_time_sizer_copy, 1, wx.EXPAND, 0)
        sizer_4.Add(sizer_7, 2, wx.EXPAND, 0)
        filter_VehType_sizer_copy.Add(self.list_VehTypes_filter_dlg, 1, wx.ALL | wx.EXPAND, 5)
        sizer_4.Add(filter_VehType_sizer_copy, 2, wx.EXPAND, 0)
        sizer_8.Add(self.label_1_copy_1, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_4.Add(sizer_8, 1, wx.EXPAND, 0)
        sizer_10.Add(self.label_1, 0, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 5)
        from_time_sizer_copy_copy.Add(self.ith_perc, 1, wx.ALL, 5)
        sizer_7_copy.Add(from_time_sizer_copy_copy, 1, wx.EXPAND, 0)
        jth_perc.Add(self.to_time_text_dlg_copy, 1, wx.ALL, 5)
        sizer_7_copy.Add(jth_perc, 1, wx.EXPAND, 0)
        sizer_10.Add(sizer_7_copy, 3, wx.EXPAND, 0)
        sizer_11.Add(self.label_1_copy, 3, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_11.Add(self.exclusion_check_box, 1, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_10.Add(sizer_11, 2, wx.EXPAND, 0)
        sizer_11_copy.Add(self.label_1_copy_copy, 3, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_11_copy.Add(self.duplicate_check_box, 1, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 0)
        sizer_10.Add(sizer_11_copy, 2, wx.EXPAND, 0)
        sizer_11_copy_copy.Add(self.label_1_copy_copy_copy, 3, wx.LEFT | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_11_copy_copy.Add(self.stopover_textbox, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_10.Add(sizer_11_copy_copy, 2, wx.EXPAND, 0)
        sizer_4.Add(sizer_10, 4, wx.EXPAND, 0)
        sizer_3.Add(sizer_4, 6, wx.EXPAND, 0)
        sizer_6.Add(self.CancelBtn_dlg, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_6.Add(self.DLG_Process_Button, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_3.Add(sizer_6, 2, wx.EXPAND, 0)
        self.SetSizer(sizer_3)
        self.Layout()
        # end wxGlade
        
    def Cancel_Click(self, event): # wxGlade: ImportMtxDialog.<event_handler>
        self.Destroy()
    
    def Get_Exclusions(self):
        """
        Procedura wylaczajaca pewne relacje.
        Dla kazdej pary kodow CL z pliku exclusions sciezki zawierajace te pare zostaja przeciete w srodku tej pary. 
        Np. tablica GH 234 byla wykryta w CL: 2,31,53,85.
        Wowczas jesli w pliku exclusions wystepuje linia "31;53", to w bazie danych zostana zapisane dwie podroze: 2,31 i 53,85.
        """
        '''
        try:
            file = open(self.Parent.Paths["Exclusions"])
        except:
            return []
        CLsExcludeList = []    
        line_ind = 0
        lines = len(file.readlines())
        file = open(self.Parent.Paths["Exclusions"])
        while 1:
            line_ind += 1
            line = file.readline()            
            if not line:
                break
            lline = list(line)
            try:
                ind = lline.index(';')
            except:
                self.Parent.ErrMsg("Bad exclusion.txt file format")
                return []
            CLsExcludeList.append([int(line[0:ind]), int(line[ind + 1:])])
        '''
        CLsExcludeList = [] 
        res=self.Parent.DB.cur.execute('select FromCLCode,ToCLCode from Matrix where enabled=="no"').fetchall()
        for r in res:
            CLsExcludeList.append([str(r[0]),str(r[1])])
                
        return CLsExcludeList  

    def Calc_Click(self, event): # wxGlade: ImportMtxDialog.<event_handler>
        # DONE SN SPRAWDZIC, CZY WYNIKI SA POPRAWNE. 
        # DONE SN SPRAWDZIC, CZY WSZYSTKIE FILTRY DZIALAJA POPRAWNIE (PERCENTYLE, VEHTYPE, CZAS). 
        # DONE SN SPRAWDZIC, CZY DETECTED JEST TAKIE JAK POWINNO BYC
        # DONE SN SPRAWDZIC CZY NALICZAJA SIE ERRORS
        # SN: KONSEKWENTNIE ZAPISYWAC ZERA JAKO NONE, ALBO JAKO "0" - teraz jest inaczej.
        def Podziel_Podroze(Res):    
            
            def slice(indexes, res):
                
                indexes.sort()
                subs = []
                if len(indexes) > 0:
                    subs = [res[:indexes[0] + 1]]
                    for i, index in enumerate(indexes):
                        if i + 1 < len(indexes):
                            subs.append(res[index + 1:indexes[i + 1] + 1])
                    subs.append(res[indexes[-1] + 1:])
                if subs == []:
                    return [res]

                return subs
            
            def get_Exclusion_indexes(Exclusions, res):
                CLs = [str(r[0]) for r in res]
                
                    
                wystapienia = []
                for Exclusion in Exclusions:
                    wystapienia_pierwszego = [j for j, x in enumerate(CLs) if x == Exclusion[0]]        
                    
                    for wystapienie_pierwszego in wystapienia_pierwszego:
                        if wystapienie_pierwszego + 1 < len(CLs):
                            if CLs[wystapienie_pierwszego + 1] == Exclusion[1]:
                                wystapienia.append(wystapienie_pierwszego)
                wystapienia.sort()
                #if CLs[0]=='21' and CLs[-1]=='41':
                    #print res,wystapienia
                return wystapienia
            
            def get_Overtime_indexes(res, max_time_delta):
                indexes = []                
                for i, obs in enumerate(res):
                    if i < len(res) - 1:
                        delta = res[i + 1][1] - res[i][1]
                        if delta > max_time_delta:                
                            indexes.append(i)
                return indexes
            
            def TnijDuplikaty(res):
                if len(res) == 0:
                    return res
                else:
                    CLs = [r[0] for r in res]
                    podroze = [CLs]
                    iloscwyst = []
                    licz = 0
                    flaga = 0 
                    while flaga != 2:
                        flaga = 0
                        licz = licz + 1
                        for ind, pod in enumerate(podroze):
                            for indi, ppod in enumerate(pod):
                                if pod.count(ppod) > 1:
                                    flaga = 1
                                    indp1 = pod.index(ppod)
                                    indp2 = pod.index(ppod, indi + 1)
                                    przyklad = slice([indp2 - 1], pod)
                                    podroze.__delitem__(ind)
                                    podroze.insert(ind, przyklad[1])
                                    podroze.insert(ind, przyklad[0])
                
                                    break
                                elif ind == len(podroze) - 1 and indi == len(pod) - 1 and flaga == 0:
                                    flaga = 2
                    index = []
                    for ind, ppod in enumerate(podroze):
                        index.append(len(ppod))
                        
                    cum_sum = []
                    y = 0
                    for i in index:   # <--- i will contain elements (not indices) from n
                        y += i   # <--- so you need to add i, not n[i]
                        cum_sum.append(y - 1)   
                        
                    podroze1 = slice(cum_sum, res)
                    podroze1.__delitem__(-1)
                
                    return podroze1
            
            
            
            
            
            Res_ = []
            
            for res in Res:       
                Overtime_indexes = get_Overtime_indexes(res, self.stopover_threshold)
                subs = slice(Overtime_indexes, res)
                for sub in subs:
                    
                    Res_.append(sub)
            Res = Res_
            
            if self.exclude:
                Exclusions = self.Exclusions
                #print Exclusions
                Res_ = []
                for res in Res:
                    Exclusion_indexes = get_Exclusion_indexes(Exclusions, res)
                    #if Exclusion_indexes:
                        #print Exclusion_indexes
                    subs = slice(Exclusion_indexes, res)
                    #if Exclusion_indexes:
                        #print subs
                    for sub in subs:
                        Res_.append(sub)
                    #if Exclusion_indexes:
                        #print Res_
                Res = Res_
            
            if self.duplicates:
                Res_ = []
                for res in Res:    
                    subs = TnijDuplikaty(res)
                    for sub in subs:
                        Res_.append(sub)        
                Res = Res_    
            
            return Res
            
        def Process_Single_CLs(): 
            self.dialog = wx.ProgressDialog ('Progress', "Processing Single CLs", maximum=len(CLs))
            u = 0      
            
            for CL in CLs:
                self.dialog.Update(u)
                u += 1                
                Query = "SELECT PLATENO FROM DETECTEDVEHICLESVOL where CLCODE='" + CL + "'"
                Filter_Result = self.Parent.DB.con.execute(Query).fetchall()              
                pl = [a[0] for a in Filter_Result]
                un = pl.count('-')                
                l = str(len(Filter_Result))                
                Query = "UPDATE CountLocations SET VOL = '" + l + "', VOL_ERROR = '" + str(un) + "' where CLCODE='" + CL + "'"              
                self.Parent.DB.con.execute(Query)
                Query = "UPDATE Matrix SET APNR_VOLUME_ERROR = '" + str(un) + "' where FromCLCode= '" + CL + "' and ToCLCode= '" + CL + "'"
                self.Parent.DB.con.execute(Query)
                Query = "UPDATE Matrix SET APNR_VOLUME_DETECTED = '" + l + "' where FromCLCode= '" + CL + "' and ToCLCode= '" + CL + "'"
                self.Parent.DB.con.execute(Query)
                Query = "UPDATE Matrix SET APNR_VOLUME_ANY = '" + l + "' where FromCLCode= '" + CL + "' and ToCLCode= '" + CL + "'"
                self.Parent.DB.con.execute(Query)

            self.Parent.DB.con.commit()
            self.dialog.Destroy()
        
        def Process_DB(CL_dict, FromTime=None, ToTime=None, VehType=None):
            
            '''
                CL1        CL2        CL3        CL4        CL5
                |          |          |          |          |
                x========>x==========>x==========>x========>x
            
                Let's assume that within selected time interval vehicle with plate no HH-DS 2343 
                was detected on count locations, as shown above
                APNR_VOLUME_OD:
                trip is regarded as Cl1-CL5
                
                Volume_OD_Detected:
                trip is split into adjacent CL pairs: 
                CL1-CL2, CL2-CL3, CL3-CL4, CL4-CL5
                
                Volume_OD_Any:
                trip is split into acceptable "sub-trips":
                CL1-CL2,CL1-CL3,CL1-CL4,CL1-CL5,CL2-CL3,CL3-CL4,CL3-CL5,CL4,CL5
                
                Those matrices might be confusing, but we left them as they might be helpful in route choice analysis.

            '''
            
            def get_info(res):
                fromCL = str(res[0][0])
                toCL = str(res[-1][0])
                
            
                    
                allCL = [str(r[0]) for r in res]
                alltimes = [r[1] for r in res]

                travel_times = [res[i + 1][1] - res[i][1] for i, r in enumerate(res[:-1])]
                travel_time = res[-1][1] - res[0][1]
                
                #OD
                tmin,tmax=self.Parent.DB.cur.execute('select mint,maxt from Matrix where FromClCode =? and ToClCode= ?',(fromCL,toCL)).fetchall()[0]
                veh_type = res[0][2]
                a = []

                ind1 = CL_dict[fromCL]
                ind2 = CL_dict[toCL]
                if travel_time>tmin and travel_time<tmax:
                    a = list(APNR_VOLUME_OD[ind1][ind2])
                    a.append([alltimes, allCL, veh_type])
                    APNR_VOLUME_OD[ind1][ind2] = a
                #DETECTED
                for i, r in enumerate (alltimes[:-1]):
                    ind1 = CL_dict[allCL[i]]
                    ind2 = CL_dict[allCL[i + 1]]
                    trtime = alltimes[i + 1] - alltimes[i]
                    tmin,tmax=self.Parent.DB.cur.execute('select mint,maxt from Matrix where FromClCode =? and ToClCode= ?',(allCL[i],allCL[i + 1])).fetchall()[0]
                    if trtime>tmin and trtime<tmax:
                        b = list(APNR_VOLUME_DETECTED[ind1][ind2])
                        b.append([trtime, veh_type])
                        APNR_VOLUME_DETECTED[ind1][ind2] = b 
                
                #ANY
                for i in range (len(alltimes)-1):
                    for j in range(i,len(alltimes)):
                        ind1 = CL_dict[allCL[i]]
                        ind2 = CL_dict[allCL[j]]
                        trtime = alltimes[j] - alltimes[i]
                        tmin,tmax=self.Parent.DB.cur.execute('select mint,maxt from Matrix where FromClCode =? and ToClCode= ?',(allCL[i],allCL[j])).fetchall()[0]
                        if trtime>tmin and trtime<tmax:
                            b = list(APNR_VOLUME_ANY[ind1][ind2])
                            b.append([trtime, veh_type])                                                        
                            APNR_VOLUME_ANY[ind1][ind2] = b
                                                    
                            

            try:
                self.Parent.DB.cur.execute('''drop table DetectedVehiclesVol''')
            except:
                pass
            
            self.Parent.DB.cur.execute("""create table DetectedVehiclesVol(IdD INTEGER PRIMARY KEY, 
                                                CLCode INT, 
                                                DetectionTime INT,
                                                DetectionTimeIP INT, 
                                                VehType VARCHAR, 
                                                PlateNo)""")
     
            if (VehType == None or VehType == 'None'):
                operator_VT = ' <> '
                var_VT = 'any'
            else:
                operator_VT = ' = '
                var_VT = VehType
            
            self.Parent.DB.cur.execute('Insert into DetectedVehiclesVol(CLCode,DetectionTime,DetectionTimeIP,VehType,PlateNo) select CLCode,DetectionTime,DetectionTimeIP,VehType,PlateNo from DetectedVehicles where (DetectionTime between ? and ?) and VehType ' + operator_VT + ' ? ', (FromTime, ToTime, var_VT))
            self.Parent.DB.con.commit()
            
            self.Parent.DB.cur.execute("SELECT DISTINCT PLATENO FROM DETECTEDVEHICLESVOL where PlateNo <> '-' AND PLATENO <>''")
            PlateNos = self.Parent.DB.cur.fetchall()            
            

            self.dialog = wx.ProgressDialog ('Progress', "Processing Database ", len(PlateNos))
            
            u = 0                                    
            for PlateNo in PlateNos:
                if PlateNo != "-":
                    u = u + 1   
                    self.dialog.Update(u)
                    Query = "SELECT CLCode ," + TIME + ",VehType FROM DETECTEDVEHICLESVOL WHERE PLATENO = '" + str(PlateNo[0]) + "' ORDER BY DETECTIONTIMEIP"
                    res = self.Parent.DB.cur.execute(Query).fetchall()
                    Res = Podziel_Podroze([res])
                    for r in Res:                        
                        get_info(r)                                
            self.dialog.Destroy()
            self.Parent.fill_grid_CLs()  
            #self.Parent.DB.cur.execute('''drop table DetectedVehiclesVol''')
        
        def Generate_OD(CL_Dict, CLs, VehType, FromTime, ToTime):
            self.dialog = wx.ProgressDialog ('Progress', "Generating OD matrices 1/3", len(APNR_VOLUME_OD))
            u = 0
            adjust = True
            
            for i, Row in enumerate(APNR_VOLUME_OD):
                u += 1
                self.dialog.Update(u)                
                for j, Col in enumerate(Row):
                    med = 0                 
                    A = list(APNR_VOLUME_OD[i][j])
                    Vol = str(len(A))
                    if len(A) == 0:
                        A = []
                        Times = [0]
                    else:                        
                        Times = [a[0][-1] - a[0][0] for a in A if len(a[0]) > 1]
                        if len(Times) == 0:
                            Times = [0]
                        elif adjust:                            
                            m = percentile(Times, int(self.ith_perc.GetValue()))
                            M = percentile(Times, int(self.to_time_text_dlg_copy.GetValue()))
                            Times = [t for t in Times if (t >= m and t <= M)]                            
                            Vol = str(len(Times))
                            med = median(Times)
                            if len(Times) == 0:
                                Times = [0]
                    fromCL = CLs[i]
                    toCL = CLs[j]
                    Query = "UPDATE Matrix SET APNR_VOLUME_OD = '" + Vol + "', APNR_TMIN_OD  = '" + str(min(Times)) + "', APNR_TMEAN_OD  = '" + str(numpy.mean(Times)) + "', APNR_TMOD_OD  = '" + str(med) + "', APNR_TMAX_OD  = '" + str(max(Times)) + "' where FromCLCode='" + fromCL + "' and ToCLCode='" + toCL + "'"            
                    self.Parent.DB.con.execute(Query)
            self.dialog.Destroy()
            
            self.dialog = wx.ProgressDialog ('Progress', "Generating OD matrices 2/3", len(APNR_VOLUME_OD))
            u = 0
                        
            for i, Row in enumerate(APNR_VOLUME_DETECTED):
                u += 1
                self.dialog.Update(u)
                for j, Col in enumerate(Row): 
                    if i != j:
                        med = 0                 
                        A = list(APNR_VOLUME_DETECTED[i][j])
                        Vol = str(len(A))
                        if len(A) == 0:
                            A = []
                            Times = [0]
                        else:                        
                            Times = [a[0] for a in A]
                            if len(Times) == 0:
                                Times = [0]
                            elif adjust:  
                             
                                m = percentile(Times, int(self.ith_perc.GetValue()))
                                M = percentile(Times, int(self.to_time_text_dlg_copy.GetValue()))
                                Times = [t for t in Times if (t >= m and t <= M)]                            
                                Vol = str(len(Times))
                                med = median(Times)
                                if len(Times) == 0:
                                    Times = [0]
                        fromCL = CLs[i]
                        toCL = CLs[j]
                        Query = "UPDATE Matrix SET APNR_VOLUME_DETECTED = '" + Vol + "', APNR_TMIN_DETECTED  = '" + str(min(Times)) + "', APNR_TMEAN_DETECTED  = '" + str(numpy.mean(Times)) + "', APNR_TMOD_DETECTED  = '" + str(med) + "', APNR_TMAX_DETECTED  = '" + str(max(Times)) + "' where FromCLCode='" + str(fromCL) + "' and ToCLCode='" + str(toCL) + "'"            
                        self.Parent.DB.con.execute(Query)
            
            self.dialog.Destroy()
            self.dialog = wx.ProgressDialog ('Progress', "Generating OD matrices 3/3", len(APNR_VOLUME_ANY))
            u = 0
            
            for i, Row in enumerate(APNR_VOLUME_OD):
                u += 1
                self.dialog.Update(u)
                for j, Col in enumerate(Row): 
                    if i != j:
                        med = 0                 
                        A = list(APNR_VOLUME_ANY[i][j])
                        Vol = str(len(A))
                        
                        if len(A) == 0:
                            A = []
                            Times = [0]
                        else:                        
                            Times = [a[0] for a in A]
                            if len(Times) == 0:
                                Times = [0]
                            elif adjust:  
                             
                                m = percentile(Times, int(self.ith_perc.GetValue()))
                                M = percentile(Times, int(self.to_time_text_dlg_copy.GetValue()))
                                Times = [t for t in Times if (t >= m and t <= M)]                            
                                Vol = str(len(Times))
                                med = median(Times)
                                if len(Times) == 0:
                                    Times = [0]
                        fromCL = CLs[i]
                        toCL = CLs[j]
                        Query = "UPDATE Matrix SET APNR_VOLUME_ANY = '" + Vol + "', APNR_TMIN_ANY  = '" + str(min(Times)) + "', APNR_TMEAN_ANY  = '" + str(numpy.mean(Times)) + "', APNR_TMOD_ANY  = '" + str(med) + "', APNR_TMAX_ANY  = '" + str(max(Times)) + "' where FromCLCode='" + str(fromCL) + "' and ToCLCode='" + str(toCL) + "'"            
                        self.Parent.DB.con.execute(Query)
                        
            Process_Single_CLs()            
            self.dialog.Destroy()
            
                    
        ''' Get VehTypes'''
        selections = [a for a in self.list_VehTypes_filter_dlg.GetSelections()]
        filter_VehTypes = [str(self.list_VehTypes_filter_dlg.GetString(selection)) for selection in selections]
        if 'None' in filter_VehTypes and len(filter_VehTypes) > 1:
            self.Parent.ErrMsg("None as a value for VehType filter can be selected only as a single selection.")
            return
        if len(filter_VehTypes) == 1: filter_VehTypes = filter_VehTypes[0]
        if filter_VehTypes == 'Any': filter_VehTypes = None        
        
        '''get FromTime ToTime'''
        filter_FromTime = self.from_time_text_dlg.GetValue()
        filter_ToTime = self.to_time_text_dlg.GetValue()
        
        filter_FromTime = self.Parent.hh__sec(filter_FromTime)
        filter_ToTime = self.Parent.hh__sec(filter_ToTime)
        if filter_FromTime == 999999999999:
            return
        if filter_ToTime == 999999999999:
            return
        """get process trips params"""        
        self.stopover_threshold = int(self.stopover_textbox.GetValue())*60
        self.duplicates = self.duplicate_check_box.GetValue()
        self.exclude = self.exclusion_check_box.GetValue()
        if self.exclude:
            self.Exclusions = self.Get_Exclusions()
        
        if self.Parent.Interpolate:
            TIME = "DetectionTimeIP"
        else:
            caption = "Mind that using not interpolated detection times \ncan generate unstable results."
            dlg = wx.MessageDialog(self, caption, "i2 APNR", wx.YES_NO | wx.ICON_QUESTION)
            result = dlg.ShowModal() == wx.ID_YES
            dlg.Destroy()
            if not result:
                return
            TIME = "DetectionTime"
                
        self.Parent.DB.cur.execute("Select Distinct CLCODE from CountLocations")
        CLs = self.Parent.DB.cur.fetchall()
        
        CL_dict = {}        
        for i, CL in enumerate(CLs):
            CL_dict[str(CL[0])] = i

        APNR_VOLUME_OD = [[[] for i in CLs] for j in CLs]
        APNR_VOLUME_DETECTED = [[[] for i in CLs] for j in CLs]
        APNR_VOLUME_ANY = [[[] for i in CLs] for j in CLs]
        
        CLs = [CL[0] for CL in CLs]      
        
        Process_DB(CL_dict, filter_FromTime, filter_ToTime, filter_VehTypes)
        
        Generate_OD(CL_dict, CLs, filter_VehTypes, filter_FromTime, filter_ToTime)
        self.Parent.DB.con.commit()
        self.Destroy()
        #TO DO SN: po wykonaniu zapelnic jescze raz filtr na pierwszym panelu CLs
        #TO DO SN: Wyrzuc statystyki na panel matrix - jaki filtr, jaki czas, etc.


class ImportMtxDialog(wx.Dialog):
    def __init__(self, *args, **kwds):
        # begin wxGlade: ImportMtxDialog.__init__
        kwds["style"] = wx.DEFAULT_DIALOG_STYLE
        wx.Dialog.__init__(self, *args, **kwds)
        self.txt_1 = wx.StaticText(self, -1, "Import SkimMtx number:")
        self.text_ctrl_1 = wx.TextCtrl(self, -1, "")
        self.button_1 = wx.Button(self, -1, "Import")

        self.__set_properties()
        self.__do_layout()
        
        self.Bind(wx.EVT_BUTTON, self.__handler_import, self.button_1)
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: ImportMtxDialog.__set_properties
        self.SetTitle("Import SkimMtx for "+self.Parent.MinMax)
        self.SetSize((336, 80))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: ImportMtxDialog.__do_layout
        sizer_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(self.txt_1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_1.Add(self.text_ctrl_1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizer_1.Add(self.button_1, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.SetSizer(sizer_1)
        self.Layout()
        # end wxGlade
    
    def __handler_import(self,event):        
        Mtx=self.Parent.Visum.Net.Matrices.ItemByKey(self.text_ctrl_1.GetValue()).GetValuesDouble()        
        Nos=self.Parent.Visum.Net.Zones.GetMultiAttValues("No")        
        for FromCL,row in enumerate(Mtx):
            for ToCL,el in enumerate(row):                                           
                Query = "UPDATE Matrix SET " + self.Parent.MinMax+ " = '" + str(el) + "' where FromCLCode= '" + str(int(Nos[FromCL][1])) + "' and ToCLCode= '" + str(int(Nos[ToCL][1])) + "'"
                self.Parent.DB.con.execute(Query)        
        self.Parent.DB.con.commit()
        self.Parent.handler_filtrujPth(self)
        self.Destroy()
        
def VisumInit(path=None):
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
    ####
    VISUM INIT
    """
    import win32com.client        
    Visum = win32com.client.Dispatch('Visum.Visum.125')
    if path != None: Visum.LoadVersion(path)
    return Visum



