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
9. Parts of this code cannot be used to any other software creating without written permission of i2

March 2012, Krakow Poland
"""
import wx
import APNR as APNR_module

class APNR_GUI(APNR_module.APNR_GUI):
    def __init__(self,Visum):
        APNR_module.APNR_GUI.__init__(self,Visum)

stand_alone=True
try:
    Visum
    stand_alone=0
except:
    Visum=APNR_module.VisumInit("C:/small.ver")

if __name__ == "__main__":
    if stand_alone:
        app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    APNR = APNR_GUI(Visum)
    app.SetTopWindow(APNR)
    APNR.Show()
    if stand_alone:
        app.MainLoop()
