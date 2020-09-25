 ## Failsafe Development             : Bhaumik Patel,Ganapati S, Siddhesh Sahane, Uday Nikam, Sayali Mendhe
## Failsafe Development             : Bhaumik Patel,Ganapati S, Siddhesh Sahane, Uday Nikam, Sayali Mendhe
## Devloped Date                    : 04/04/2017
## Purpose                          : Failsafe Automation
## Message_Counter Devlopment       : Tanmay Bhopi,Bhaumik Patel,Sayali Mendhe,Ankit Khandelwal
## Devloped Date                    : 04/01/2017
## Purpose                          : Message Counter Automation
## GateWay Devlopment               : Tanmay Bhopi,Bhaumik Patel,Sayali Mendhe,Ankit Khandelwal
## Devloped Date                    : 04/01/2017
## Purpose                          : GateWay Automation
## Canape GateWay Devlopment        : Rajkumar , Vishwamber,bhaumik, Zulfikar, Sachin, Shalini
## Devloped Date                    : 28/06/2017
## Purpose                          : GateWay Automation
## Error Handling And Stabilizing   : Harish Y, Uday Nikam, Rahul Vaity, Sayali Mendhe
## Devloped Date                    : Till Date
## Purpose                          : to make script error free
 ########################## All includes and modules to be imported #######################
 ########################## All includes and modules to be imported #######################
from __future__ import with_statement
import Tkinter

from Tkinter import *
import ttk 
from ttk import Progressbar, Style
from ordereddict import OrderedDict
##from collections import OrderedDict 
import tkFileDialog
import os
import time
import datetime
import xlrd
import xlwt
import xlutils
import math
from xlutils.copy import copy
import  simplejson as json , uuid
##import  json , uuid
from pprint import pprint as pprint
import threading
from threading import Thread
import logging 
from Tkinter import Tk, Frame, BOTH, RIGHT, RAISED 
from PIL import ImageTk , Image
import ImageGrab
import rtplib                                                                                                           # Some utilities to determine the platform type and the path of the used experiment
import platformmanager, cdacon, os, sys, dSPACEDemoUtilities
from cdautomationlib  import *
import tkMessageBox
import ctypes
import shutil
import win32com.client
import subprocess
import win32con
import win32api
import win32gui
import win32process
import pythoncom
import distutils.dir_util
import win32com.client
import errno
import re
import csv
##import copy
import unicodedata
import glob
from xlrd import open_workbook
import re, traceback
from time import sleep


##from distutils import msvc9compiler
##import psutil 

##########################################################################################
global myAppl
global Missing_Input_Details
Missing_Input_Details = ""
global Script_Path, Org_Path
global error_Count
error_Count = 0
Script_Path = os.getcwd()                                                                                       # Getting current directory of the script
print "Current directoty is ",Script_Path
pathindex = Script_Path.rindex("\\")
global All_Applications,All_Applications_Result,All_Applications_copy
All_Applications = []
All_Applications_copy = []
All_Applications_Result = []
Script_Path1 = Script_Path[:pathindex]                                                                                   # One directory back 
print "Base path is ",Script_Path1
pathindex = Script_Path1.rindex("\\")
global Test_Sheet_Path
Org_Path = Script_Path1[:pathindex+1]                                                                                   # Two directories back
print "Base path is ",Org_Path
Image_Path = Script_Path1 + "\\" + "IMAGES" + "\\"
print "Image_Path is ",Image_Path
##########################  Initialization of Global Variables ###########################]
Platform_Name=''
Browse_button_color='red'
start_button_color =''
testNo_cnt_Final=''
Experiment=''
DispatchSheet=''
Data= ''
VehicleName= ''
RegionName= ''
FilePath1 = Org_Path
FilePath2 = Org_Path
Variant = ''
App_Arry_Final= ''
Variant_Test_Enabled = ''
PartNo = ''
Ances_array=[]                                                                                                          # Intialization of global variables
tree=''
Browse_button_color = ''
tree_frame=''
scrollbar = ''
uid= 0
curItem=0
end_test_str = 'end_test_case'
start_test_str = 'test_case_no'
type_col = 2
no_testcase_col = 2
temp = 0
row_num = 0
Destination_Folder_Path = ' '
Wait_Over = 0
TestCase_End_Row = 0
TestCase_Start_Row = 0
start_test_case_list = []
end_test_case_list = []
VariantName_tree = ''
ApplicationName_tree = ''
TestCaseName_tree = ''
Test_Sheet_Path = Org_Path + '03_Master_Test_Sheet'
Var = 'Variant'
AdasECU = ''
No_Rows = 0
ratio = 0
Failsafe_Enabled = 0
sig_data_sheet_str = 'Signal_Data'
myAppl = None
pfm = None
Variant_Path = ' '
Variant_Value = []
prev_uid = 0
cur_uid = 0
Dest_Folder_Path_Vehicle = ''
LayoutConfig = ''
LayoutDiag = ''
LayoutMeter = ''
LayoutSide = ''
LayoutEap= ''
LayoutMrr= ''
LayoutFrC= ''
LayoutSow= ''
Power_Supply_path = ''
CAR_SLCT_NO_path = ''
frame1_color = '#%02x%02x%02x' % (196, 190, 180)
browse_frame_color = '#%02x%02x%02x' % (175, 175, 175)
default_button_color = '#%02x%02x%02x' % (146, 208, 80)
val=0
 
            
def Variant_write(myAppl,Write_Var):
    try:
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(1)
        if AdasECU == 'Dual ADAS':
            myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS2/ADAS2_EXIST/Value').Write(1)
            myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS3/ADAS3_EXIST/Value').Write(1)
            myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)         # Write Variant change value to variable         
            time.sleep(1)
            print "DIAG_CMD_NO_path at beginning", DIAG_CMD_NO_path
            myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C3_TX/DIAG_CMD_NO/Value').Write(4)
            time.sleep(.5)
            myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C3_TX/DIAG_CMD_NO/Value').Write(0)
            time.sleep(1)
            myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID75D_TX/DIAG_CMD_NO/Value').Write(4)
            time.sleep(.5)
            myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID75D_TX/DIAG_CMD_NO/Value').Write(0)

        elif AdasECU == 'Dual ADAS2 Not Exist':
            myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS2/ADAS2_EXIST/Value').Write(0)
            myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS3/ADAS3_EXIST/Value').Write(1)
            myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)         # Write Variant change value to variable
            time.sleep(1)
            myAppl.Variable(DIAG_CMD_NO_path).Write(4)
            time.sleep(.5)
            myAppl.Variable(DIAG_CMD_NO_path).Write(0)


        elif AdasECU == 'Dual ADAS3 Not Exist':
            myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS3/ADAS3_EXIST/Value').Write(0)
            myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS2/ADAS2_EXIST/Value').Write(1)
            myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)         # Write Variant change value to variable
            time.sleep(1)
            myAppl.Variable(DIAG_CMD_NO_path).Write(4)
            time.sleep(.5)
            myAppl.Variable(DIAG_CMD_NO_path).Write(0)
        
        else :
            myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)         # Write Variant change value to variable
            time.sleep(1)
            myAppl.Variable(DIAG_CMD_NO_path).Write(4)
            time.sleep(.5)
            myAppl.Variable(DIAG_CMD_NO_path).Write(0)



        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(129)

    except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')

            logging.exception('Power_Supply_path,DIAG_CMD_NO_path or CAR_SLCT_NO_path not found')

# After OK button in MAIN tab of GUI is pressed this function is excuted. It disables all buttons except start button and thus freezes Checkbutton Values# 
def OK_Pressed():
    Plantmodel_button["state"] = DISABLED
    dispatch_button["state"] = DISABLED
    dispatch_button["state"] = DISABLED
    start_button["state"] = NORMAL
    stop_button["state"] = DISABLED
    reset_button["state"] = DISABLED
    
    ALL_check_button["state"] = DISABLED
    ITS_check_button["state"] = DISABLED
    FAILSAFE_check_button["state"] = DISABLED
    GATEWAY_check_button["state"] = DISABLED
    BUSOFF_check_button["state"] = DISABLED
    ACTIVE_check_button["state"] = DISABLED
    MSG_COUNTER_check_button["state"] = DISABLED
    CONFIG_CHECK_check_button["state"] = DISABLED
    ICC_Cancel_Testing_button["state"] = DISABLED
    OK_button["state"] = DISABLED
    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info('You have selected the following \n  %s',Enabled_Validation_Items)
    print "OK"
     

        
def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))


    
        
    
def Assign():

    global ALL_enabled,ITS_enabled,Message_Counter_enabled_GUI,DDT_enabled_GUI,Failsafe_Enabled_GUI,BusOff_Enabled_GUI,Gateway_DIAG_Enabled_GUI,Gateway_TGW_Enabled_GUI,Config_check_Enabled_GUI,Active_Enabled_GUI,ICC_CANCEL_Enabled_GUI
    global Enabled_Validation_Items
    ALL_enabled=0
    ITS_enabled=0
    Message_Counter_enabled_GUI=0
    Config_check_Enabled_GUI=0
    DDT_enable_GUI=0
    Failsafe_Enabled_GUI=0
    BusOff_Enabled_GUI=0
    Gateway_DIAG_Enabled_GUI=0
    Gateway_TGW_Enabled_GUI=0
    Active_Enabled_GUI=0
    ICC_CANCEL_Enabled_GUI=0
    Enabled_Validation_Items=[]

    def ACTIVE_TREE():
        ACTIVE_TREE=construct_JSON_tree(ActiveTestDict,frame5)

    if v1.get() == 1 :
        ALL_enabled=1
    
        
    if v2.get() == 1:
        ITS_enabled = 1
        Enabled_Validation_Items.append("ITS Testing")
      

    if v3.get() == 1:
        DDT_enabled_GUI = DDT_enabled and 1
        Enabled_Validation_Items.append("DDT Testing ")
        

    if v4.get() == 1:

        Message_Counter_enabled_GUI = Message_Counter_enabled and  1
        Enabled_Validation_Items.append("Message Counter Testing")
        
    if v5.get() == 1:

        Failsafe_Enabled_GUI = Failsafe_Enabled and 1
        
        Enabled_Validation_Items.append("FailSafe Testing")
        
    if v6.get() == 1:
        Gateway_DIAG_Enabled_GUI =  Gateway_DIAG_Enabled and 1
        Gateway_TGW_Enabled_GUI = Gateway_TGW_Enabled and 1
        Enabled_Validation_Items.append("GW/TGW Testing")
        
    if v7.get() == 1:
        BusOff_Enabled_GUI =  BusOff_Enabled_GUI and 1
        Enabled_Validation_Items.append("BusOff Testing")
        
    if v8.get()	== 1:
        Config_check_Enabled_GUI =  Config_Check_Enabled and 1
        Enabled_Validation_Items.append("Config_check Testing")

    if v9.get()	== 1:
        ICC_CANCEL_Enabled_GUI =  ICC_Cancel_Check_Enabled and 1
        print "ICC_CANCEL_Enabled_GUI",ICC_CANCEL_Enabled_GUI,ICC_Cancel_Check_Enabled
        Enabled_Validation_Items.append("ICC_CANCEL_Enabled_GUI")        
        
    if v10.get()	== 1:
        Active_Enabled_GUI =  Config_Check_Enabled and 1
        Enabled_Validation_Items.append("Active Testing")
                

    if ( v1.get()==0 and v2.get()==0 and v3.get()==0 and v4.get()==0 and v5.get()==0 and v6.get()==0 and v7.get()==0 and v8.get()== 0 ):
        print ("No Button is Pressed")
        Enabled_Validation_Items.append("No Button is Pressed")


    OK_button["state"] = NORMAL   

#***********************************************#
                
#***********Resest button Functionality*********#   
       
def Reset():
    global PlatformName,platformmanager
    Plantmodel_entrybox.delete(0, END)
    dispatch_entrybox.delete(0, END)
    ##tree.destroy()
    ##scrollbar.destroy()
    shutil.rmtree(Dest_Folder_Path_Vehicle)
    time.sleep(3)
    Instrumentation().AnimationMode =0                                                                      # Exit button functionality to close control desk and python
    time.sleep(4)
    PlatformManager().Platforms.Item(PlatformName).Stop()
    time.sleep(5)
    os.system('TASKKILL /F /IM ControlDesk.exe')
    time.sleep(5)
    os.rmdir(Dest_Folder_Path_Vehicle)
    reset_button["state"] = DISABLED                                                                            # Disable Reset button
    dispatch_button["state"] = NORMAL                                                                           # Enable Dispatch Sheet browse button
    Plantmodel_button["state"] = DISABLED                                                                       # Disable Plant model browse button
    start_button["state"] = DISABLED                                                                            # Disable Start button
    stop_button["state"] = DISABLED                                                                             # Disable Stop button

#***********************************************#
#******** Stop Button Functionality ************#

def Stop():
    stop_button["state"] = DISABLED
    start_thread.terminate()                                                                                    # Stop button disabled after single press
    time.sleep(1)
    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info('Test case execution stopped')
    start_button["state"] = NORMAL
    reset_button["state"] = NORMAL

    Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)

#***********************************************#

def copy_folder(src, dest):
    try:
        shutil.copytree(src, dest)
    except OSError, e:
        if e.errno == errno.ENOTDIR:
            shutil.copy(src, dest)
        else:
            print("Directory not copied Error: %s" %e)
  


def Summary_Sheet_Function() :
    summary_tool_path = Script_Path + "\\" +  "Summary_Sheet_Filling_Tool.xls"
                    
    xlapp1 = win32com.client.Dispatch("Excel.Application")   #To open Excel 

    if os.path.exists(str(summary_tool_path)):

        xlapp1.Workbooks.Open(Filename=str(summary_tool_path), ReadOnly=1)
                
        xlapp1.Application.Run("Summary_Sheet_Filling_Tool.xls!module1.Report_Creation")

    xlapp1.Workbooks.Close()

    
    time.sleep(5)


def Controldesk_Load_Reload(Write_Var):
    global myAppl,SystemType,CAR_SLCT_NO_path
    print CAR_SLCT_NO_path,"CAR_SLCT_NO_path",Write_Var
    myAppl = None
    Instrumentation().AnimationMode =0
    time.sleep(2)
    PlatformManager().Platforms.Item(PlatformName).Stop()
    time.sleep(4)
    myAppl = rtplib.Appl(ApplFile + ".sdf", PlatformName, SystemType)
    time.sleep(4)
    PlatformManager().Platforms.Item(PlatformName).Start()
    time.sleep(6)
    Instrumentation().AnimationMode =2
    time.sleep(2)

    Variant_write(myAppl,Write_Var)
    
    myAppl.Variable(Power_Supply_path).Write(1)                                                                
    time.sleep(5)

          
#***********Start button Functionality**********#

def Start():
    global curItem, Ances_array, Log_File, VariantName_tree,\
           ApplicationName_tree, TestCaseName_tree, end_test_str,\
           start_test_str, No_Rows , ratio, TestCase_Start_Row,\
           TestCase_End_Row,  sig_data_sheet_str, LayoutConfig ,\
           AdasECU,WorkBook10, folder_name_dispatch, Wait_Over ,A2l_Path, \
           Master_Result_Report_FailSafe,CA_Dest_Folder, \
           Failsafe_Delete_row,BUSOFF_JUDGEMENT_SHEET,Message_Counter_Report_path,Gateway_Diag_Report_path,Gateway_TGW_Report_path, Fails,Failsafe_Enabled_GUI,Failsafe_Testing_ADAS25,ICC_CANCEL_Enabled_GUI
    global book,sig_data_sheet,Test_Sheet_Path
    global judge_sheet_path,Book_Master_TP
    #failsafe_cat__drop["state"] = DISABLED

    global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DIAG_CMD_NO_path_3,DTC_string_temp,DTC_string_1_temp,DTC_string_path_temp,Read_vehicle_speed_path
    #GLOBAL SHEET PATHS# 
    global interface_sheet_path,Test_Sheet
  
    

    judge_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Master_Judgement_Sheet_ITS.xls'      #MASTER JUDGEMENT SHEET
  

    #INTERFACE SHEET
    interface_sheet_path = Org_Path + "\\" + "02_Script" + "\\" + "JUDGEMENT_SCRIPTS" + "\\" + "Interface_VBA.xls"
  


    #Message Counter Sheet
    Message_Counter_Report_path = Org_Path + "05_Master_Result_Reports" + "\\" + "Message_Counter_Report"  + "\\Message_Counter_Report.xlsx"
    InterfaceTextFile = Org_Path + "04_Master_CANape_Configurations/03_MESSAGE_COUNTER/Interface.txt"
    CAN_MSG_DATA_PATH=  Org_Path + "\\" + "02_Script" + "\\" + "JUDGEMENT_SCRIPTS" + "\\" +"Message_Counter_Judgement_Sheet.xls"

    #Gateway Sheet
    Gateway_Diag_Report_path = Org_Path + "05_Master_Result_Reports" + "\\" + "Gateway_Diag_TGW_Report"  + "\\Gateway_Diag_TGW_Report.xls"
    Gateway_TGW_Report_path = Org_Path + "05_Master_Result_Reports" + "\\" + "Gateway_TGW_Report"  + "\\Gateway_TGW_Report.xlsx"
    Gateway_Trace_Path= Org_Path + "\\" + "02_Script" + "\\" + "JUDGEMENT_SCRIPTS" + "\\" + "Gateway_Judgement.xls"
    
    #Busoff Sheet
    BUSOFF_JUDGEMENT_SHEET=  Org_Path + "\\" + "02_Script" + "\\" + "JUDGEMENT_SCRIPTS" + "\\" +"Busoff_Judgement_Sheet.xls"

    # ADAS 3/3D Failsafe Sheet
    Master_Result_Report_FailSafe = Org_Path + "05_Master_Result_Reports"+ "\\" + "Failsafe_Result_Report" + "\\"  + "Master_Result_Report_FailSafe.xls"
    Master_Result_FailSafe_Path = Org_Path + "05_Master_Result_Reports" + "\\" + "Failsafe_Result_Report" + "\\"  + "Master_Result_Report_FailSafe.xls"

     # ADAS 2.5 Failsafe Sheet
    Master_Result_Report = Org_Path + "05_Master_Result_Reports"


    #Config Check
    Config_check_report_folder_path = Org_Path + '05_Master_Result_Reports' +"\\"+ 'Config_check'

    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
    logging.info('Plant model selected %s %s',VehicleName)    


            
            

    #****** These are paths to various files and folders used for Message Counter script  **********#
   
    Master_CANape_Configuration_Path=  Org_Path + "04_Master_CANape_Configurations"
    Master_CANape_Message_Counter_Path =  Master_CANape_Configuration_Path +"\\"+"03_MESSAGE_COUNTER"  # Path of MASTER CANape files of Message Counter
    Master_CANape_BusOff_Path =           Master_CANape_Configuration_Path + "\\"+ "06_Busoff_Check"  # Path of MASTER CANape files of Busoff
    Master_CANape_FailSafe_Path =         Master_CANape_Configuration_Path+ "\\" + "05_Failsafe"
    Master_CANape_ADAS25_Failsafe_Path =  Master_CANape_Configuration_Path + "\\05_Failsafe_CANape_Configuration"+"\\01_Failsafe_reference_folder"
    
    
    A2l_Path = Org_Path + "05_Master_Result_Reports" + "\\" + "Failsafe_Result_Report" + "\\"  + "adas1.a2l"
   
   


    #****** These are paths to various files and folders used for Gateway script  **********#
    Master_CANape_Gateway_DIAG_Path = Org_Path + "04_Master_CANape_Configurations" +"\\"+"04_GATEWAY"+"\\"+"01_Gateway_DIAG"  # Path of MASTER CANape files of Message Counter
    Master_CANape_Gateway_TGW_Path = Org_Path + "04_Master_CANape_Configurations" +"\\"+"04_GATEWAY"+"\\"+"02_Gateway_TGW"  # Path of MASTER CANape files of Message Counter

    
    Config_check_report_folder_path = Org_Path + '05_Master_Result_Reports' +"\\"+ 'Config_check'                              #Config_check report path
    pythoncom.CoInitialize()
    curItem = 0	    				
    SignalData = []
    SigInfo = []
    SigNames = []
    execute_cont_str = ['exec_cont']
    Var_Val = 0
    sheetNumber = 0
    exec_var_dep = ['exec_var_dep']
    exec_delay = ['exec_delay']
    exec_push_var_dep = ['exec_push_var_dep']
    exec_wait_var_dep = ['exec_wait_var_dep']
    execute_start_end_str = ['exec_start_end'] 
    sig_info={}
    Log_File_Path = Script_Path +'\\HILS_Testing_Log.txt'
    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
    logging.info('Test case execution started')                                                                 # Logging info in the log file
    Summary_Sheet_Path = 'Result_Summary_Report'
    CA_Sheet_Path = 'Caution_Advisory_Report'
    Summary_src_Folder = Org_Path + '05_Master_Result_Reports' + "\\" +  Summary_Sheet_Path
    CA_Sheet_Folder = Org_Path + '05_Master_Result_Reports' + "\\" +  CA_Sheet_Path
    Summary_dest_Folder = Dest_Folder_Path_Vehicle
    CA_Dest_Folder = Dest_Folder_Path_Vehicle + '\\02_CautionAdvisory'
    CA_Test_Procedure_Folder = Org_Path + '05_Master_Result_Reports' + "\\"  + 'Failsafe_Result_Report' #GAnapathi 18th april
    Dispatch_src_Folder  = FilePath1
    print 'Dest_Folder_Path_Vehicle ' + Dest_Folder_Path_Vehicle
    Dispatch_dest_Folder = Dest_Folder_Path_Vehicle + '\\' + folder_name_dispatch
    print 'Dispatch_src_Folder' + Dispatch_src_Folder
    print 'Dispatch_dest_Folder' + Dispatch_dest_Folder
    distutils.dir_util.copy_tree(Dispatch_src_Folder,Dispatch_dest_Folder)
    distutils.dir_util.copy_tree(Summary_src_Folder,Summary_dest_Folder)
    distutils.dir_util.copy_tree(CA_Sheet_Folder,CA_Dest_Folder)




 
   
#************ADAS_HILS_AUTOMATION***************#
    
    def ADAS_HILS_AUTOMATION (SignalData,SigInfo,sig_data_sheet ,myAppl):
        try:
            print "ADAS_HILS_AUTOMATION"
            wait_confirm_val = 50000
##            myAppl.Variable(Power_Supply_path).Write(1)
##            time.sleep(.5)
##            myAppl.Variable("simState").Write(0)                                                                    # 'Reset' Simstate
##            time.sleep(.5)
##            myAppl.Variable("simState").Write(2)                                                                    # 'Set' Simstate
##            time.sleep(2)
            
            
            
         
            for m in range(len(SignalData)):                    # Loop for execution type "exec_start_end" and "exec_cont" 
                Signal_Data = []
                Signal_Data = SignalData[m]
               
                test_sig_type = Signal_Data[0]
                
                
                ##ITS_progressbar["value"] = 6
                if test_sig_type in ['0']:
                    sig_name =  Signal_Data[1]                                                                      # Collect signal name
                    sig_path = Signal_Data[2]                                                                       # Collect signal path  
                    sig_delay = float(Signal_Data[3])        # Collect the delay value . This delay will be executed after desired value for the signal is set 
                    sig_val = Signal_Data[4]                 # Collect the various values to be set for a particular signal. Note that 'sig_val' is a 'list' 

                    for n in range(len(sig_val)):                        
                        myAppl.Variable(sig_path).Write(sig_val[n])
                    
                        if sig_delay < 0:    # If delay specified is less than zero, this loop will confirm if the desired value is being set to the signal. Else specified delay is executed.
                            temp_count = 0 
                            while temp_count < wait_confirm_val:
                                temp_count = temp_count + 1                                
                                temp_val = None
                                temp_val = myAppl.Variable(sig_path).Read()                                         # Value written to the signal is read
                                time.sleep(0.5)
                                if temp_val == sig_val[n]:                                                          # After confirmation loop ends
                                    break
                        else:
                            time.sleep(float(sig_delay))
                    
                elif test_sig_type in ['1']:                    
                    sig_name = Signal_Data[1]                                                                       # Collect the name of the signal to which value is to be set        
                    sig_path = Signal_Data[2]                                                                       # Collect the path of the above signal
                    sig_val = Signal_Data[3]                                                                        # Collect the value to be set to the above signal
                    dep_var_name = Signal_Data[4]                                                                   # Collect the name of the dependent variable.  
                    dep_var_path = Signal_Data[5]                                                                   # Collect the path of the dependent variable
                    dep_var_cond = Signal_Data[6]                                                                   # Collect the dependency condition
                    dep_var_val = float(Signal_Data[7])                                                             # Collect the value of dependent variable 

                                                                                                                    
                    if dep_var_cond in ['>']:                                                                       # Execute this loop if dependency condition is "greater than"
                        temp_count = 0
                        

                        while temp_count < wait_confirm_val:
                           
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                           # print "temp_val",temp_val
                           # print "dep_var_val",dep_var_val
                            if temp_val > dep_var_val:                                                              # After the condition is met value is written to the signal
                                print "temp_val",temp_val
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Write the desired value to the signal only if dependent variable meets the condition with specified dependent value
                                break

                                                            
                    if dep_var_cond in ['=']:                                                                       # Execute this loop if dependency condition is "equal to" 
                        temp_count = 0
                        
                        while temp_count < wait_confirm_val:
                            
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))                              # Value of the dependent variable is read


                            if temp_val == dep_var_val:                                                             # After the condition is met value is written to the signal
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Write the desired value to the signal only if dependent variable meets the condition with specified dependent value
                                break

                                             
                        
                    if dep_var_cond in ['<']:                                                                       # Execute this loop if dependency condition is "less than" 
                        temp_count = 0
                       
                        while temp_count < wait_confirm_val:
                           
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))                              # Value of the dependent variable is read

                            if temp_val < dep_var_val:                                                              # After the condition is met value is written to the signal
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Write the desired value to the signal only if dependent variable meets the condition with specified dependent value
                                break
                            
                elif test_sig_type in ['2']:
                    sig_name = Signal_Data[1]
                    sig_delay = float(Signal_Data[2])
                 #   print "\n waiting for delay = " + str(sig_delay)
                    time.sleep(sig_delay)

           
                
                elif test_sig_type in ['3']:
                    sig_name = Signal_Data[1]                                                                       # Collect the name of the signal to which value is to be set        
                    sig_path = Signal_Data[2]                                                                       # Collect the path of the above signal
                    sig_val = Signal_Data[3]                                                                        # Collect the value to be set to the above signal
                    dep_var_name = Signal_Data[4]                                                                   # Collect the name of the dependent variable.  
                    dep_var_path = Signal_Data[5]                                                                   # Collect the path of the dependent variable
                    dep_var_cond = Signal_Data[6]                                                                   # Collect the dependency condition
                    dep_var_val = float(Signal_Data[7])                                                             # Collect the value of dependent variable
              

                                                                 
                
                    if dep_var_cond in ['=']:                                                                       # Execute this loop if dependency condition is "equal to" 
                        temp_count = 0
                        while temp_count < wait_confirm_val:
                            time.sleep(0.5)
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            temp_count = temp_count + 1                           
                            
                            if temp_val == dep_var_val:                                                             # After the condition is met....value is written to the signal 
                               break
                            else:                                
                                myAppl.Variable(sig_path).Write(0)                                                  # Writing the value '0' is push button feature
                                time.sleep(0.1)
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Writing the desired value to the signal
                                time.sleep(0.1) 
                                myAppl.Variable(sig_path).Write(0)

                                                   
              
                
                    if dep_var_cond in ['>']:                                                                       # Execute this loop if dependency condition is "greater than" 
                        temp_count = 0
                        while temp_count < wait_confirm_val:
                            time.sleep(0.5)
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            temp_count = temp_count + 1                           
                            
                            if temp_val > dep_var_val:                                                              # After the condition is met....value is written to the signal 
                               break
                            else:                                
                                myAppl.Variable(sig_path).Write(0)                                                  # Writing the value '0' is push button feature
                                time.sleep(0.1)
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Writing the desired value to the signal
                                time.sleep(0.1) 
                                myAppl.Variable(sig_path).Write(0)

             
                
                    if dep_var_cond in ['<']:                                                                       # Execute this loop if dependency condition is "less than" ##
                        temp_count = 0
                        while temp_count < wait_confirm_val:
                            time.sleep(0.5)
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            temp_count = temp_count + 1                           
                            
                            if temp_val < dep_var_val:                                                              # After the condition is met....value is written to the signal 
                               break
                            else:                                
                                myAppl.Variable(sig_path).Write(0)                                                  # Writing the value '0' is push button feature
                                time.sleep(0.1)
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Writing the desired value to the signal
                                time.sleep(0.1) 
                                myAppl.Variable(sig_path).Write(0)


            
                                
                elif test_sig_type in ['4']:                    
                    sig_name = Signal_Data[1]                    
                    dep_var_name = Signal_Data[2]
                    dep_var_path = Signal_Data[3]
                    dep_var_cond = Signal_Data[4]
                    dep_var_val = float(Signal_Data[5])
                
                    if dep_var_cond in ['>']:
                        
                        while(1): 
                            
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))
                            time.sleep(0.2)
                            
                            if temp_val > dep_var_val:
                               break
                        
                    elif dep_var_cond in ['<']:
                        
                        while temp_count < wait_confirm_val:
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val = myAppl.Variable(dep_var_path).Read()
                            time.sleep(0.2)                          
                            if temp_val < dep_var_val:
                               # myAppl.Variable(sig_path).Write(sig_val)
                                break
                        
                    elif dep_var_cond in ['=']:

                        while(1): 
                            
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))
                            time.sleep(0.2)
                            if temp_val == dep_var_val:
                                break
                elif test_sig_type in ['5']:                    
                                    
                    dep_var_name = Signal_Data[2]
                    sig_path = Signal_Data[3]
                    dep_var_val = float(Signal_Data[5])

                    dep_var_name = dep_var_name + ";"
                    temp_sig_path = sig_path
                    flag = 0
                    i = 0
                    old_Rx_status = 9
                    Tx_status = 16
                    while ((len(dep_var_name)) >= 2):
                        sig_path = temp_sig_path
                        
                        if i == len(dep_var_name):
                             break;   
                        if i == 0:
                           sig_path = sig_path + "1/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                           
                        elif i == 1:
                           sig_path = sig_path + "2/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                           
                        elif i == 2:
                           sig_path = sig_path + "3/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 3:
                           sig_path = sig_path + "4/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 4:
                           sig_path = sig_path + "5/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 5:
                           sig_path = sig_path + "6/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 6:
                           sig_path = sig_path + "7/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 7:
                           sig_path = sig_path + "8/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                           temp_dep_var_name = ''
                           for j in range(i+1,len(dep_var_name)):
                               temp_dep_var_name = temp_dep_var_name + dep_var_name[j]
                           dep_var_name = temp_dep_var_name
                           i = -1
                        
                           time.sleep(1)
                           while(1):

                                  Rx_status = myAppl.Variable(Test_Automation_path)\
                                              .Read()
                                  time.sleep(0.01)
                                
                                  if Rx_status > old_Rx_status:
                                       old_Rx_status = Rx_status
                                       myAppl.Variable(Start_CANape_path)\
                                       .Write(Tx_status)
                                       Tx_status = Tx_status + 1
                                       break;
                              

                time.sleep(0.5)
                sig_path = None                                                                                     # Set default values to all the signals present in sheet2 of excel sheet
                sig_name = None
                sig_value = None
                sig_reset = None
                ##ITS_progressbar["value"] = 8

                time.sleep(0.5)
                #myAppl.Variable(Power_Supply_path).Write(129)
        except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.exception('Signal_paths are wrong')
   
#***********************************************#

#********************Exit_func******************#  

    def  Exit_func():
        
        time.sleep(3)
        Instrumentation().AnimationMode =0                                                                      # Exit button functionality to close control desk and python
        time.sleep(4)
        PlatformManager().Platforms.Item(PlatformName).Stop()
        time.sleep(5)
        os.system('TASKKILL /F /IM ControlDesk.exe')
        time.sleep(5)
        os.system('TASKKILL /F /IM pythonwin.exe')
        
        #os.system('TASKKILL /F /IM pythonwin.exe')
        
#***********************************************#

#********************Exit_func******************#  

    def  Close_Control_Desk_func():
        Log_File_Path = Script_Path +'\\HILS_Testing_Log.txt'
        time.sleep(3)
        Instrumentation().AnimationMode =0                                                                      # Exit button functionality to close control desk and python
        time.sleep(4)
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                            format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Animation Mode Deactivated') 
        PlatformManager().Platforms.Item(PlatformName).Stop()
        time.sleep(5)
        
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                            format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Plant Model Deloaded') 
        os.system('TASKKILL /F /IM ControlDesk.exe')
        time.sleep(5)

        
#***********************************************#
        
        
#******get_data_execute_cont function***********#            

    def get_data_execute_cont(sig_name,sig_path, Application_sheet,
                              sig_delay,x, sig_data_sheet):
        exec_type = '0'                                                                                         # Define execution type as '0'
        delay_col = 2
        val = [(float(str(Application_sheet.cell(x,delay_col+1).value)))]                                         # Get value to be set to the signal in the list 'val'
        str_val = str(Application_sheet.cell(x,delay_col+1).value)                                              # Get the value to be set to the signal in string format 
        count = 2
        #********#

        if (Application_sheet.cell(x,delay_col+ count).value) != '':
            str_val = str(Application_sheet.cell(x,delay_col+count).value)
            count= count+1                                  
            temp = float(str_val)                            
            val.append(temp)                                                                                    # If cell is not blank, the value in the cell gets appended in the list 'val'
        signal_data = [exec_type,sig_name,sig_path,sig_delay,val]                                               # Collect all the signal data like path, execution type, delay value ....etc
        
        return signal_data
#***********************************************#

    def get_data_execute_Failsafe_cont(sig_name,sig_path, Application_sheet,
                              sig_delay,x, sig_data_sheet):
        exec_type = '0'                                                                                         # Define execution type as '0'
        Application_col = 7

        val = [(float(str(Application_sheet.cell(x,Application_col+1).value)))]                                         # Get value to be set to the signal in the list 'val'
        str_val = str(Application_sheet.cell(x,Application_col+1).value)                                              # Get the value to be set to the signal in string format 
        count = 2

        #********#
##        print "Application_sheet.cell(x,Application_col+ count).value)"
##        print Application_sheet.cell(x,Application_col+count).value
##        if (Application_sheet.cell(x,Application_col+ count).value) != '':
##            str_val = str(Application_sheet.cell(x,Application_col+count).value)
##            count= count+1                                  
##            temp = float(str_val)                            
##            val.append(temp)                                                                                    # If cell is not blank, the value in the cell gets appended in the list 'val'
        signal_data = [exec_type,sig_name,sig_path,sig_delay,val]                                               # Collect all the signal data like path, execution type, delay value ....etc
    
        return signal_data



#*****get_data_execute_var_dep function*********#

    def get_data_execute_var_dep(sig_name,sig_path, Application_sheet,
                                 sig_delay,x, sig_data_sheet):
        exec_type = '1'                                                                                         # Define execution type as '1try:
        delay_col = 2
        dep_var = str(Application_sheet.cell(x,delay_col).value)                                                # Get dependent variable name from 'delay_col'
        for m in range(0,sig_data_sheet.nrows ):
            if dep_var == sig_data_sheet.cell(m, 0).value:
                sig_path_dep_var = sig_data_sheet.cell(m, 1).value
                
        dep_var_chk = str(Application_sheet.cell(x,delay_col+1).value)                                          # Get the dependent condition
        dep_var_val = float(str(Application_sheet.cell(x,delay_col+2).value))                                   # Get the value to be met by the dependent variable according to the dependent condition
        val = float(str(Application_sheet.cell(x,delay_col+3).value))                                           # Get the value to be set to the signal in column1.
                                
        signal_data = [exec_type,sig_name,sig_path,val,
                       dep_var,sig_path_dep_var,
                       dep_var_chk,dep_var_val]                                                                 # Collect all the signal data like path, execution type, delay value 

        return signal_data
#***********************************************#
    
#**********get_execute_delay function***********#


    def get_execute_delay(sig_name,sig_delay):
        exec_type = '2'                                                                                         # Define execution type as '2'     
        #time.sleep(sig_delay)
        signal_data = [exec_type,sig_name,sig_delay]                                                            # Get the value to be set to the signal in column1.
        return signal_data

#***********************************************#
    

    

#***********************************************#

#************Start CANape function**************#
    
    def Start_CANape(proj_path):
        All_Process_TM = os.popen("tasklist").read()
        while "canape32.exe" in All_Process_TM:
            All_Process_TM = os.popen("tasklist").read()
            os.system("taskkill /f /im canape32.exe")
            
        if os.path.exists(proj_path + "\\CANape.MDF"):
            os.remove(proj_path + "\\CANape.MDF")
        proj_path = proj_path + "\\CANape.INI"
        if os.path.isfile('C:\Program Files\Vector CANape 7.0\Exec\canape32.EXE'):
            command = ['C:\Program Files\Vector CANape 7.0\Exec\canape32.EXE',proj_path]
        else:
            command = ['C:\Program Files (x86)\Vector\CANape\\12\Exec\canape32.exe',proj_path]
        proc1 = subprocess.Popen(command, shell =False)#, stderr=subprocess.PIPE)
        hwndMain = 0
        while (hwndMain == 0):
            hwndMain = win32gui.FindWindow(None,"DISCLAIMER")
        print "hwndMain",hwndMain
        time.sleep(0.5)
        try:           
            win32gui.SetForegroundWindow(hwndMain)   
            win32api.PostMessage(hwndMain,win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            
        except:
            print "hwndMain",hwndMain
            win32api.PostMessage(hwndMain,win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    
    def ASCII_to_CSV(ascfile_name,csvfile_name):
  
        num_lines=0
        old_TS=0
        MIN_ROW=0
        MAX_ROW=0
        num_lines2=0
      
        with open(ascfile_name,'r') as c:
                for line1 in c:
                        num_lines2 =num_lines2 + 1
                        if "Length" not in line1:
                                continue
                        sep = 'Length'
                        line1 = line1.split(sep, 1)[0]
                        line1 = re.sub(r'[. \n]+', " ", line1);
                        Time_Stamp1=line1[1:10]
                        Time_Stamp1=Time_Stamp1.replace(' ', '.')
                        TS=  line1[1:3]
                        if(TS=='14' and old_TS=='13'):
                                MIN_ROW = num_lines2
                        if(TS=='15' and old_TS=='14'):
                                MAX_ROW = num_lines2
                        old_TS=TS

        with open(csvfile_name,'ab') as csvfile:

                with open(ascfile_name) as f:

                        for line in f:
                                num_lines =num_lines + 1
                              
                                if(num_lines > MIN_ROW and num_lines < MAX_ROW):
                                        
                                        if "Length" not in line:
                                                continue
                                        
                                        csv_data=[]
                                        sep = 'Length'
                                        
                                        line = line.split(sep, 1)[0]
                                        #line = re.sub(r'[. \n]+', " ", line);

                                        Time_Stamp=line[2:11]
                                        Time_Stamp=Time_Stamp.replace(' ', '.')                                
                                        Channel = line[12]
                                        CANID = line[15:18]
                                        TxRx = line[31:33]
                                        DLC = line[38]
                                        BINARY_DATA=str(line[40:-1])
               
                                        csv_data.append(str(Time_Stamp))
                                        csv_data.append(Channel)
                                        csv_data.append(CANID)
                                        csv_data.append(TxRx)
                                        csv_data.append(DLC)
                                        csv_data.append(BINARY_DATA)
                
                                        data = csv.writer(csvfile,delimiter=",",quotechar="'")
                                        data.writerows([csv_data])
                                        
                #data.writerow(" ")
                #data.writerow(" ")        
    def Exit_Button_Function() :
        exit_start=Tk()
        w1 = 280                                                                                                # Width of the application window
        h1 = 50                                                                                                 # Height of the applicaiton window
        sw1 = exit_start.winfo_screenwidth()                                                                    # Width of the screen
        sh1 = exit_start.winfo_screenheight()                                                                   # Height of the screen
        x1 = (sw1 - w1)/2                                                                                       # X co ordinate
        y1 = (sh1 - h1)/2                                                                                       # Y co ordinate
        exit_start.geometry('%dx%d+%d+%d' % (w1, h1, x1, y1))            
        exit_frame = Frame(exit_start, relief = RAISED,
                           borderwidth=2, bg = frame1_color)
        exit_frame.place(x = 0, y = 0)
        exit_label= Label(exit_frame, justify="center",
                          text ='HILS Testing execution Completed',
                          bg =heading_color, fg = "black",
                          font="Times 10 bold")
        exit_label.pack(padx = 5, pady =4, side =RIGHT)
        exit_button = Tkinter.Button(exit_frame, text = "Exit",                                                 # Exit frame functionality
                                     bg = Browse_button_color,
                                     activebackground = "red",
                                     height = 2, width = 8,
                                     relief = RAISED, bd = 3,
                                     cursor = "hand2",
                                     command= Exit_func)
        exit_button.pack(padx = 5, pady =4, side = LEFT)
        exit_start.mainloop




    
    def ITS_Application_Testing():
        global Var_Val, myAppl
        global summary_tool_path,Test_Sheet_Path,judge_sheet_path
        global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DTC_string_path_temp,Read_vehicle_speed_path
        global Book_Master_TP

        print "Book_Master_TP",Book_Master_TP        
        curItem = 0	    				
        SignalData = []
        SigInfo = []
        SigNames = []
        execute_cont_str = ['exec_cont']
        Var_Val = 0
        sheetNumber = 0
        exec_var_dep = ['exec_var_dep']
        exec_delay = ['exec_delay']
        exec_push_var_dep = ['exec_push_var_dep']
        exec_wait_var_dep = ['exec_wait_var_dep']
        execute_start_end_str = ['exec_start_end'] 
        sig_info={}
        Log_File_Path = Script_Path +'\\HILS_Testing_Log.txt'
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                 							  # Creation of log file
                            format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('##############  ITS execution started  ##############')                                                                 # Logging info in the log file
         
        Var_Val=0
        end_test_str = 'end_test_case'
        start_test_str = 'test_case_no'
        global prevUID_ITS,Item
        Item=0;
        print "entered ITS"
        try:                                                        					  
            book = xlrd.open_workbook(Book_Master_TP)                                                                       					  # Opening the test case sheet
            sig_data_sheet = book.sheet_by_name(sig_data_sheet_str)
            Vehicle_Name = VehicleName + ' - ' + RegionName + ' - ' + PartNo       
            ITS_Tree=construct_JSON_tree(VehicleDict,frame3)
            ITS_vehicle_id_entry.delete(0, END)
            ITS_vehicle_id_entry.insert(0, Vehicle_Name)
            ITS_overall_progressbar["maximum"] = uid
        except:
            logging.info('Error  :: while extracting Power_Supply/CAR_SLCT_NO/DIAG_CMD_NO path')    
   
        xlapp1 = win32com.client.Dispatch("Excel.Application")   #To open Excel 

        if os.path.exists(str(interface_sheet_path)):

            xlapp1.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    
            xlapp1.Application.Run("Interface_VBA.xls!module7.Change_DBC_Name")

        xlapp1.Workbooks.Close()
        
        prevUID_ITS=uid
        curItem=0
       
        for j in range(0, uid):             
            ITS_overall_progressbar["value"] = j	
            Ances_array=[]
            TestCaseNameId = ''
            ApplicationNameId = ''
            TestCaseNameId = ''
            VariantName_tree = ''
            ApplicationName_tree = ''
            ITS_result_entry.delete(0, END)
            TestCaseName_tree = ''
            ITS_testcase_entry.delete(0, END)
            ITS_application_entry.delete(0, END)
            ITS_variant_entry.delete(0, END)
            curItem= curItem + 1
            ITS_Tree.selection_set(curItem)
            Item=curItem
            try:
                Master_CANape_Path = Org_Path + '04_Master_CANape_Configurations\\01_ITS_Testing_CANape_Confurigations'    # Creating path for source folder to copy CANape files from Master folder  
                book1 = xlrd.open_workbook(judge_sheet_path,formatting_info=True)
                sheetNames = book1.sheet_names()
                sheetNumber = 0
                for i in range(0, 10):
                    src_folder = Master_CANape_Path
                    
                    dest_folder = folder_path                                                                   # Path for destination folder to copy CANape files from Master folder
                    
                    ParentItem=ITS_Tree.parent(Item)
                    Ances_array.append(ParentItem)
                    Item=Ances_array[i]
                  
                    if Ances_array[i]=='':
                        break
                    else :
                        continue
                     
                   
                Heir = len(Ances_array)
              
                ITS_progressbar["maximum"] =  Heir                                                                      # Progressbar max length assigned
                if Heir == 6:
                    

                    TestCaseNameId = curItem
                    ApplicationNameId = Ances_array[0]
                    VariantNameId = Ances_array[1]
                    
                elif Heir == 5:
                   
                    ApplicationNameId = curItem
                    VariantNameId = Ances_array[0]
                    TestCaseNameId = ''
                    
                elif Heir == 4:
                    VariantNameId = curItem
                    ApplicationNameId = ''
                    TestCaseNameId = ''
                else:
                    VariantNameId = ''
                    ApplicationNameId = ''
                   
                    TestCaseNameId = ''
            
                if VariantNameId=='':
                    ITS_variant_entry.delete(0, END)
                else:
                    VariantName_tree = ITS_Tree.item(VariantNameId, 'text')
                    print "VariantName_tree",VariantName_tree
                    data  = VariantName_tree.split('_')
                    VariantName_tree = ' Variant ' + data[2]
                    ITS_variant_entry["state"] = NORMAL
                    ITS_variant_entry.delete(0, END)
                    ITS_variant_entry.insert(0, VariantName_tree)
                    print 'present variant is ' + VariantName_tree
                    ITS_variant_entry["state"] = DISABLED
                    
                if ApplicationNameId=='':
                    ITS_application_entry.delete(0, END)
                else:
                    ApplicationName_tree = ITS_Tree.item(ApplicationNameId, 'text')
                    ITS_application_entry["state"] = NORMAL
                    ITS_application_entry.delete(0, END)
                    print 'present appli is ' + ApplicationName_tree
                    ITS_application_entry.insert(0, ApplicationName_tree)
                    ITS_application_entry["state"] = DISABLED
                  
                if TestCaseNameId == '':
                    ITS_testcase_entry.delete(0, END)
                else:
                    TestCaseName_tree = ITS_Tree.item(TestCaseNameId, 'text')
                    ITS_testcase_entry["state"] = NORMAL
                    ITS_testcase_entry.delete(0, END)
                    ITS_testcase_entry.insert(0, TestCaseName_tree)
                    print 'present test case is ' + TestCaseName_tree
                    ITS_testcase_entry["state"] = DISABLED
                    Log_File_Path = Script_Path +'\\HILS_Testing_Log.txt'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Test case is - %s',TestCaseName_tree)
                        
                if TestCaseName_tree == '' and ApplicationName_tree == '' and \
                   VariantName_tree == '':
                    print ' While Vehicle features '   
                    time.sleep(1)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Vehicle is - %s',Vehicle_Id)
                    
                    continue
               
                elif TestCaseName_tree == ''  and ApplicationName_tree == ''and\
                     VariantName_tree!= '' :
                    print 'Var_Val', Var_Val
                    try:
                        os.makedirs(dest_folder) 
                    except:
                        logging.info("Info","Directory Exist")
                        
                    myAppl.Variable(Power_Supply_path).Write(1)
                    Write_Var = Variant_Value[Var_Val]
                    print " Write_Var: ", Write_Var
                    Variant_write(myAppl,Write_Var)
                    Var_Val = Var_Val + 1
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Variant is - %s', VariantName_tree) 
                    time.sleep(3)
                        
                    continue
                
                elif TestCaseName_tree == ''  and ApplicationName_tree != '' and VariantName_tree!= '':
             
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Application is - %s', ApplicationName_tree)
                    Sub_VariantName_tree = data[2]
                    print "Sub_VariantName_tree",Sub_VariantName_tree
                    Report_src_folder = Org_Path + '05_Master_Result_Reports' + "\\" +  ApplicationName_tree
                    print 'Report_src_folder' + Report_src_folder
                    print  'dest_folder' + dest_folder
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    if not '_' in Sub_VariantName_tree :
                        
                        Sub_VariantName_tree = '_' + Sub_VariantName_tree
    ####                        
                    
                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break
                    
                    dest_folder = dest_folder +'\\' +  folder_name_app
                    print "folder_name_app",folder_name_app
                    try:
                        os.makedirs(dest_folder) 
                    except:
                        logging.info("Info","Directory Exist")
                        
    ##                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
    ##                    for directories_d in dird:
    ##                        if ApplicationName_tree == directories_d[3:]:
    ##                            dest_folder = dest_folder + '\\' + directories_d
    ##                            print "ApplicationName_tree dest_folder",dest_folder
    ##                            break
                    print len(All_Applications),"len(All_Applications)"
                    All_Applications_copy = []
                    for i in range(0,len(All_Applications)):
                        
                        if i < 10:
                            ApplicationName_tree_copy = str("0")+ str(i+1) + "_" + All_Applications[i]
                        else:
                            ApplicationName_tree_copy = str(i+1) + "_" + All_Applications[i]
                        if ApplicationName_tree in ApplicationName_tree_copy:
                            ApplicationName_tree_folder = ApplicationName_tree_copy
                            print "ApplicationName_tree_folder",ApplicationName_tree_folder
                        All_Applications_copy.append(ApplicationName_tree_copy)
                        
                    print "All_Applications_copy",All_Applications_copy
                    if ApplicationName_tree_folder in All_Applications_copy:
                        
                        dest_folder = dest_folder +'\\' + ApplicationName_tree_folder                                     #make Application folder
                    try:
                        os.makedirs(dest_folder) 
                    except:
                        logging.info("Info","Directory Exist")

                    screenshot_dest_folder = dest_folder +'\\' + "00_Screenshots"                                       #make 00_Screenshots folder
                    try:
                        os.makedirs(screenshot_dest_folder) 
                    except:
                        logging.info("Info","Directory Exist")
                                            
                        
                    Report_dest_folder = dest_folder
                    print 'Report_dest_folder ' + Report_dest_folder
                    distutils.dir_util.copy_tree(Report_src_folder,Report_dest_folder)
                    rep_name = ApplicationName_tree + '.xls'
                    Rep_Name = ApplicationName_tree  + Sub_VariantName_tree+ '.xls'
                    os.rename(os.path.join(Report_dest_folder,rep_name),os.path.join(Report_dest_folder,Rep_Name))
                    Rep_Path = os.path.join(Report_dest_folder,Rep_Name)
                    print Rep_Path
    ##                    interface_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Interface_VBA.xls'
                    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet 

                    if os.path.exists(str(interface_sheet_path)):

                        xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    
                        xlapp.Application.Run("Interface_VBA.xls!module5.Hide_All",Rep_Path)


                    xlapp.Workbooks.Close()
                    
                    print ApplicationName_tree
                    Application_sheet = book.sheet_by_name(ApplicationName_tree)
                    End_row = Application_sheet.nrows
                    Data_End_Row =  sig_data_sheet.nrows                  


                    time.sleep(1)

                    continue
                
                else:
                    flag = 0
                    Sub_TestCaseName_tree = TestCaseName_tree[-2:]
                    ssdest_folder = ''
                    print "Sub_TestCaseName_tree",Sub_TestCaseName_tree
                    print 'Src_Folder' + src_folder


                    print "ApplicationName_tree",ApplicationName_tree
                    book1 = xlrd.open_workbook(judge_sheet_path,formatting_info=True)
                    sheetNames = book1.sheet_names()
                    sheetNumber = 0
                    for i in sheetNames:
                        if "Py-MScript" in i:
                            break
                        sheetNumber = sheetNumber + 1
                    book2 = copy(book1)
                    Interface_sheet = book2.get_sheet(sheetNumber)
                    Interface_sheet.write(0, 1, VehicleName)
                    print "VehicleName",VehicleName
                    Interface_sheet.write(1, 1, RegionName)
                    print "RegionName",RegionName
                    Interface_sheet.write(2, 1, Write_Var)
                    Interface_sheet.write(3, 1, ApplicationName_tree)
                    print "ApplicationName_tree",ApplicationName_tree

                    Interface_sheet.write(4, 1, Sub_TestCaseName_tree)
                    Interface_sheet.write(22, 1, 0)     # Failsafe = 0 to run ITS in (B23 cell)
                    book2.save(judge_sheet_path)

                    ITS_CANape_src_folder = src_folder +'\\01_ITS_reference_folder' 
      
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
    ####                        Sub_VariantName_tree = '_' + Sub_VariantName_tree
                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break
                    dest_folder = dest_folder +'\\' +  folder_name_app       
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    for directories_d in dird:
                        if ApplicationName_tree == directories_d[3:]:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "ApplicationName_tree dest_folder",dest_folder
                            ssdest_folder = dest_folder
                            break
                    dest_folder = dest_folder + '\\' + directories_d + '_' + Sub_TestCaseName_tree
                    try:
                        os.mkdir(dest_folder)
                    except:
                        pass
                    if (flag == 1):
                        print "flag in if ",flag
                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.info("Test case not found in Master Configuration %s" + TestCaseName_tree + "  %s" + Sub_TestCaseName_tree)
                    else:
                     
                        distutils.dir_util.copy_tree(ITS_CANape_src_folder,dest_folder)
                    
                    
    ##                    interface_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Interface_VBA.xls'
                    xlapp = win32com.client.Dispatch("Excel.Application")   #To open Excel 

                    if os.path.exists(str(interface_sheet_path)):

                        xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    
                        xlapp.Application.Run("Interface_VBA.xls!module6.Make_Canape")
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                            level=logging.INFO,
                                            format='%(message)s')
                    logging.info('CANape Log copy paste done')

                    xlapp.Workbooks.Close()

                    write_ss_info = open(dest_folder + "\\Screen_Shot.txt","w")
                    write_ss_info.write(ssdest_folder + "\\" + "00_Screenshots\n")
                    write_ss_info.write(TestCaseName_tree[-2:])
                    write_ss_info.write("_")
                    write_ss_info.write(ApplicationName_tree)
                    write_ss_info.close()

                    
                    Application_sheet = book.sheet_by_name(ApplicationName_tree)

                # Actual execution of test case    
                ITS_progressbar["maximum"] = 10
                End_row = Application_sheet.nrows
                Data_End_Row =  sig_data_sheet.nrows
                TestS = TestCaseName_tree[-2:]
                end_test_str = end_test_str[:13] + '_' + TestS
                start_test_str = start_test_str[:12] + '_' + TestS
               
                for k in range(0, End_row):
                    
                    if Application_sheet.cell(k,0).value== start_test_str :
                        TestCase_Start_Row= k
                        break
                         
                    else:
                        continue
                for k in range(0, End_row):
                    if Application_sheet.cell(k,0).value== end_test_str :
                        TestCase_End_Row= k
                        break
                    else:
                        continue
                
                
                ITS_progressbar["value"] = 2
                save_var = 0
               
                for x in range(TestCase_Start_Row + 1,TestCase_End_Row):

                    sig_name = Application_sheet.cell(x,0).value
                    sig_delay = Application_sheet.cell(x,2).value
                    for y in range(0,sig_data_sheet.nrows ):
                        if sig_name == sig_data_sheet.cell(y, 0).value:
                            sig_path = sig_data_sheet.cell(y, 1).value
                            sig_value = sig_data_sheet.cell(y, 2).value
                            sig_reset = sig_data_sheet.cell(y, 3).value
                    SigNames.append(sig_name)       
                    sig_data = [sig_path,sig_value,sig_reset]                                                       # Collect all the signal data in list 'sig_data'
                    sig_info[sig_name] = sig_data
                    ITS_progressbar["value"] = 3
                    SigInfo.append(sig_info[sig_name])
                     
                   
                    if str(Application_sheet.cell(x,1).value) in \
                       execute_start_end_str:                                                                       # If string in cloumn 2 is 'exec_start_end' then call 'get_data_execute_start_end_str' to collect test case data 
                        signal_data = get_data_execute_start_end_str(sig_name,
                                                                     sig_path,
                                                                     Application_sheet,
                                                                     sig_delay,x,
                                                                     sig_data_sheet)
                        save_var = 1
                    elif str(Application_sheet.cell(x,1).value) in execute_cont_str:                                # If string in cloumn 2 is 'exec_cont' then call 'get_data_execute_conti' to collect test case data
                        
                        print "sig_name",sig_name
                        print "sig_path",sig_path
                        
                        print "Application_sheet",Application_sheet
                        print "sig_delay",sig_delay
                        print "sig_data_sheet",sig_data_sheet
                        print "x",x
                      
                        signal_data = get_data_execute_cont(sig_name,sig_path,
                                                            Application_sheet,
                                                            sig_delay,x,
                                                            sig_data_sheet)
                        save_var = 1    
                    elif str(Application_sheet.cell(x,1).value) in exec_var_dep:                                    # If string in cloumn 2 is 'exec_var_dep' then call 'get_data_execute_var_dep' to collect test case data   
                        signal_data = get_data_execute_var_dep(sig_name,sig_path,
                                                               Application_sheet,
                                                               sig_delay,x,
                                                               sig_data_sheet)
                        save_var = 1
                    elif str(Application_sheet.cell(x,1).value) in exec_delay:                                      # If string in cloumn 2 is 'exec_delay' then call 'get_execute_delay' to collect test case data  
                        signal_data = get_execute_delay(sig_name,sig_delay)
                        save_var = 1

                    ITS_progressbar["value"] =4
                    SignalData.append(signal_data)
                    
                    if save_var == 1:                                                                               # Save_var = 1 signifies that atleast 1 signal is present in that test case 
                        save_var = 0
                        
                    signal_data = []
                    sig_info = {}
                    


               
                book1 = xlrd.open_workbook(judge_sheet_path,formatting_info=True)                                   
                JudgementsheetName = book1.sheet_by_index(sheetNumber)
                if JudgementsheetName.cell(8,1).value == "NA":                                                      #check in judgementsheet for "A" or  "NA" for testcase applicable. if testcase not applicable it will disply TBD on GUI
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Test case skipped due to test case is not Applicable in judgementsheet %s')
                    ITS_Tree.item(curItem, text = ApplicationName_tree + '_TEST_' + Sub_TestCaseName_tree, values = "TBD" )
                    continue                   
                    
                    
                pathTextFile = dest_folder
                pathTextFile = pathTextFile + "\\" + "Sync.txt"
                print "pathTextFile", pathTextFile

                

                myAppl.Variable(Power_Supply_path).Write(1)
                time.sleep(1)
                myAppl.Variable("simState").Write(0)                                                                    # 'Reset' Simstate
                time.sleep(.5)
                myAppl.Variable("simState").Write(2)                                                                    # 'Set' Simstate
                time.sleep(2)

                myAppl.Variable(Power_Supply_path).Write(129)
                time.sleep(1)
                myAppl.Variable(Power_Supply_path).Write(1)                                                                
                time.sleep(6)
                
                myAppl.Variable(DIAG_CMD_NO_path).Write(3)                                                          #Clear DTC
                time.sleep(.5)                  
                myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                time.sleep(.5)
         
                flag_speed = 0
                count = 0
                while(flag_speed == 0):                                                                             #check for vehicle speed and DTC 
                    read_velocity_1 = myAppl.Variable(Read_vehicle_speed_path).Read()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(2)                                                          #Read DTC
                    time.sleep(0.5)                 
                    Actual_DTC_set = myAppl.Variable(DTC_string_path_temp).Read()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(.5)
                    print "Actual_DTC_set",int(float(Actual_DTC_set))
                    time.sleep(7)
                    read_velocity_2 = myAppl.Variable(Read_vehicle_speed_path).Read()
                    print "vehicle speed read",int(float(read_velocity_2)),int(float(read_velocity_1))
                    if (int(float(read_velocity_2)) - int(float(read_velocity_1)) > 0) and (int(float(Actual_DTC_set)) == 0) :                    
                        flag_speed = 1
                    else:
                        myAppl.Variable(Power_Supply_path).Write(129)
                        time.sleep(1)
                        myAppl.Variable(Power_Supply_path).Write(1)                                                                 # 'Set' Simstate
                        time.sleep(6)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(3)
                        time.sleep(.5)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                        time.sleep(.5)                        
                        myAppl.Variable(DIAG_CMD_NO_path).Write(2)
                        time.sleep(0.5)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                        time.sleep(0.5)

                    if count > 1:    
                        flag_speed = 1
                    count = count + 1

                if count > 1:
                    Controldesk_Load_Reload(Write_Var)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('DTC not cleared')
                    #continue
                
                Start_CANape(dest_folder)
                time.sleep(2)

                flag_start = 0
                while(flag_start == 0):
                    
                    syncFileRead = open(pathTextFile,'r')

                    valueRead = syncFileRead.read()
                    syncFileRead.close()

                    if (valueRead == '7'):
                        time.sleep(1);
                        flag_start = 1                    
                Layout = FilePath2 + "\\" +"side.lay"
    ####                        
    ####                        print Layout
    ####                        Instrumentation().Layouts.Open(Layout)
    ##                       # Instrumentation().Layouts.Save(Layout)
    ####                        Instrumentation().ActiveLayout.Normalize()
                Instrumentation().Layouts.Item(Layout).Activate()  
                Instrumentation().ActiveLayout.Maximize()
                ##Instrumentation().ActiveLayout.Normalize()
                ADAS_HILS_AUTOMATION (SignalData,SigInfo,sig_data_sheet,myAppl)
                time.sleep(2)
                syncFileWrite = open(pathTextFile,'w')
                sync_num = 9
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)

                syncFileWrite.close()

                time.sleep(1)

                try:      
                    wildcard = ".*CANape*"
                    cW = cWindow()
                    handle_manager=cW.find_window_wildcard(wildcard)
                    cW.Maximize()
                    cW.BringToTop()
                    cW.SetAsForegroundWindow()
                    bbox = win32gui.GetWindowRect(handle_manager)

                except:
                    f = open("log.txt", "w")
                    f.write(traceback.format_exc())
                    print traceback.format_exc()                

                for m in range(1,sig_data_sheet.nrows ):

                    set_sig_reset_value = None
                    set_sig_path = None
                    set_sig_default_value = None
                    set_sig_Appl = None
                    set_sig_reset_value = int(sig_data_sheet.cell(m,3).value)
                    set_sig_path = sig_data_sheet.cell(m,1).value
                    set_sig_default_value = sig_data_sheet.cell(m,2).value
                    set_sig_Appl = sig_data_sheet.cell(m,4).value
                    if set_sig_reset_value ==1 :
                        if ((set_sig_Appl == ApplicationName_tree)|(set_sig_Appl == 'All')):
    ##                            print "if loop"
                            try:
                            
                                myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                                time.sleep(1)
                            except:
                                pass

                flag = 0
                while(flag == 0):
    ##                        print "while loop"
                    syncFileRead = open(pathTextFile,'r')

                    valueRead = syncFileRead.read()

                    syncFileRead.close()

                    if (valueRead == '8'):
                       time.sleep(3);
                       flag = 1
    ##                    Stop_CANape(dest_folder)

                interfacebook = xlrd.open_workbook(interface_sheet_path,formatting_info=True)
                interfacesheetName = interfacebook.sheet_by_name("VBA_Script_Run")
                if interfacesheetName.cell(7,1).value == 0:                                                         # if Provided judgement type not found in Result report log it to the HILS_Testing_Log file
                    print "Provided judgement type not found in Result report"
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('_________________Provided judgement type not found in Result report_________________')


                dest_files = os.listdir(dest_folder)
                for file in dest_files:
                    if file[-4:] == ".exe" or file[-4:] == ".ctf" or \
                       file[-4:] == ".cns" or file[-4:] == ".scr" or file[-2:] == ".c" or \
                       file == "Defect_Description_1.txt"  or \
                       file == "Sync.txt" :
                        os.remove(dest_folder + "\\" + file)
                try:
                    shutil.rmtree(dest_folder + "\\Test_Result_Automation_mcr")
                except:
                    pass
                try:
                    shutil.rmtree(dest_folder + "\\Screenshot_call_mcr")
                except:
                    pass

                
                book1 = xlrd.open_workbook(judge_sheet_path,formatting_info=True)
                JudgementsheetName = book1.sheet_by_index(sheetNumber)                  
                if JudgementsheetName.cell(6,1).value!=1:
                    Result_Strng = 'CA'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Result of executed test case is  %s',Result_Strng)
                else:
                    Result_Strng = 'OK'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Result of executed test case is  %s',Result_Strng)
                
                book3 = copy(book1)
                JudgementsheetName = book3.get_sheet(sheetNumber)
                JudgementsheetName.write(0, 1, '')
                JudgementsheetName.write(1, 1, '')
                JudgementsheetName.write(2, 1, '')
                JudgementsheetName.write(3, 1, '')
                JudgementsheetName.write(4, 1, '')
                JudgementsheetName.write(5, 1, '')
                JudgementsheetName.write(6, 1, '')
    ##                book3.save(judge_sheet_path)
    ##                    book3.Close()
                
                ITS_result_entry.delete(0, END)
                ITS_result_entry.insert(0,Result_Strng)
                ITS_Tree.item(curItem, text = ApplicationName_tree + '_TEST_' + Sub_TestCaseName_tree, values = Result_Strng )
                time.sleep(1)
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('Test case execution Completed %s', TestCaseName_tree)
                

                time.sleep(1)
           
                ITS_overall_progressbar["value"] = uid
                
            
                
                                
                
                #myAppl.Variable(Power_Supply_path).Write(129)                    
                

                
                ITS_progressbar["value"] = 10
                SignalData = []
                SigInfo =[]
                SigNames = []

        #pfm = None                                                                                              # Release the platformmanager
        #sig_val = None
       # temp_val = None
        #myAppl = None
                time.sleep(1)

         
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                            level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('##############  HILS Testing execution Completed  ##############')
            except Exception, e:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
                syncFileWrite = open(pathTextFile,'w')
                sync_num = 3
                ITS_Tree.item(curItem, text = ApplicationName_tree + '_TEST_' + Sub_TestCaseName_tree, values = "TBD" )
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)
                logging.exception('Test case execution stopped abrubtly')

# _________________________ICC_CANCEL_Check_Testing starts here__________________________________

    def ICC_CANCEL_Application_Testing() :
        global Var_Val, myAppl,DIAG_CMD_NO_path,CAR_SLCT_NO_path,Power_Supply_path,ICC_Cancel_code,uid
        global summary_tool_path
        global ApplicationName_tree,FilePath2,CAN_Value_1,CAN_Value_2
        global ICC_Cancel_overall_progressbar,ICC_Cancel_progressbar
        global SigInfo_temp,SignalData1_temp,SignalData,SigNames,SigInfo,sig_info
        global Covariant,Variant_Value,Test_Sheet_Path
    
        #global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DTC_string_path_temp,Read_vehicle_speed_path
        
        execute_cont_str=[]
        SigInfo_temp=[]       

        CAN_Value_1 =''
        CAN_Value_2 = ''
        folder_name_app = "10_ICC_CANCEL_Testing"                                          
        Meter_Navi_Layout_str= 'Meter_Navi_Layout_Name'
        Diag_layout_str= 'Diag_Layout_Name'
        curItem = 0	    				
        SignalData = []
        SigInfo = []
        SigNames = []
        execute_cont_str = ['exec_cont']
        Var_Val = 0
        sheetNumber = 0
        exec_var_dep = ['exec_var_dep']
        exec_delay = ['exec_delay']
        exec_push_var_dep = ['exec_push_var_dep']
        exec_wait_var_dep = ['exec_wait_var_dep']
        execute_start_end_str = ['exec_start_end'] 
        sig_info={}
        Log_File_Path = Script_Path +'\\HILS_Testing_Log.txt'
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                            format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('##############  ICC Cancel Check execution started  ##############')                                                                 # Logging info in the log file
         
        Var_Val=0
        end_test_str = 'end_test_case'
        start_test_str = 'test_case_no'
        global prevUID_ITS,Item
        Item=0;
        print "entered ICC"

        try:
            Folders = os.listdir(Test_Sheet_Path)
            print Folders
            #if Failsafe_Enabled == 1 :
           #     Master_TestPattern_name = 'Master_TestPattern_FLS.xls'
            
            Master_TestPattern_name = 'Master_TestPattern_FLS.xls'
            print "Master_TestPattern_name",Master_TestPattern_name
            if VehicleName in Folders:
                logging.info('Info :: Vehicle folder found ')
                Folders = os.listdir(Test_Sheet_Path+'\\'+VehicleName)
                print Folders
                if Master_TestPattern_name in Folders:
                    logging.info('Info :: Test pattern sheet found ')
                    Book_Master_TP = Org_Path + "03_Master_Test_Sheet" + "\\" + VehicleName+"\\" + Master_TestPattern_name
                    book_TP = xlrd.open_workbook(Book_Master_TP,formatting_info=True)													# Open Judgement sheet
                    sheetNames_TP = book_TP.sheet_names()
                    if sig_data_sheet_str in sheetNames_TP:
                        print "Signal Data sheet exist"
                        Signal_Data_sheet = book_TP.sheet_by_name(sig_data_sheet_str)
                        End_row = Signal_Data_sheet.nrows
                        print "End_row",End_row
                        Power_Supply_path_Found = 0
                        CAR_SLCT_NO_path_Found = 0
                        DIAG_CMD_NO_path_Found = 0
                        DIAG_CMD_NO_3_path_Found = 0
                        DTC_string_path_temp_Found = 0
                        DTC_String_path_Found = 0
                        DTC_String_1_path_Found = 0
                        Read_vehicle_speed_path_Found = 0
                        for k in range(0, End_row):
                            print Signal_Data_sheet.cell(k,0).value
                            if Signal_Data_sheet.cell(k,0).value == "Power_Supply":
                                Power_Supply_path_Found = 1
                                Power_Supply_path = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "CAR_SLCT_NO":
                                CAR_SLCT_NO_path_Found = 1
                                CAR_SLCT_NO_path = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "DIAG_CMD_NO_2":
                                DIAG_CMD_NO_path_Found = 1
                                DIAG_CMD_NO_path = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "DIAG_CMD_NO_3":
                                DIAG_CMD_NO_3_path_Found = 1
                                DIAG_CMD_NO_path_3 = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "DTC_String":
                                DTC_String_path_Found = 1
                                DTC_string_path_temp = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "DTC_String_1":
                                DTC_String_1_path_Found = 1
                                DTC_string_1_temp = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "DTC_string_ADAS2_ICC":
                                DTC_string_path_temp_Found = 1
                                DTC_string_path_temp_ICC = Signal_Data_sheet.cell(k, 1).value
                            elif Signal_Data_sheet.cell(k,0).value == "Read_vehicle_speed":
                                Read_vehicle_speed_path_Found = 1
                                Read_vehicle_speed_path = Signal_Data_sheet.cell(k, 1).value
                            else:
                                continue
                            
                        if Power_Supply_path_Found == 1:
                            logging.info('Info :: Power_Supply_path found in Test pattern sheet')
                        else:
                            logging.info('Info :: Power_Supply_path not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. Power_Supply_path not found in Test pattern sheet\n'
                        if CAR_SLCT_NO_path_Found == 1:
                            logging.info('Info :: CAR_SLCT_NO_path found in Test pattern sheet')
                        else:
                            logging.info('Info :: CAR_SLCT_NO_path not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. CAR_SLCT_NO_path not found in Test pattern sheet\n'
                            
                        if DIAG_CMD_NO_path_Found == 1:
                            logging.info('Info :: DIAG_CMD_NO_path_2 found in Test pattern sheet')
                        else:
                            logging.info('Info :: DIAG_CMD_NO_path_2 not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. DIAG_CMD_NO_path not found in Test pattern sheet\n'
                        if DIAG_CMD_NO_3_path_Found == 1:
                            logging.info('Info :: DIAG_CMD_NO_3_path found in Test pattern sheet')
                        else:
                            logging.info('Info :: DIAG_CMD_NO_3_path not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details +=str(error_Count) + '. DIAG_CMD_NO_3_path not found in Test pattern sheet\n'

                            
                        if DTC_String_path_Found == 1:
                            logging.info('Info :: DTC_String_path found in Test pattern sheet')
                        else:
                            logging.info('Info :: DTC_String_path not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. DTC_String_path not found in Test pattern sheet\n'
                            
##                            
##                        if DTC_String_1_path_Found == 1:
##                            logging.info('Info :: DTC_String_1_path found in Test pattern sheet')
##                        else:
##                            logging.info('Info :: DTC_String_1_path not found in Test pattern sheet')
##                            error_Count = error_Count + 1
##                            Missing_Input_Details += str(error_Count) + '. DTC_String_1_path not found in Test pattern sheet\n'
                            
                        if DTC_string_path_temp_Found == 1:
                            logging.info('Info :: DTC_string_path_temp found in Test pattern sheet')
                        else:
                            logging.info('Info :: DTC_string_path_temp not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. DTC_string_path_temp not found in Test pattern sheet\n'
                        if Read_vehicle_speed_path_Found == 1:
                            logging.info('Info :: Read_vehicle_speed_path found in Test pattern sheet')
                        else:
                            logging.info('Info :: Read_vehicle_speed_path not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. Read_vehicle_speed_path not found in Test pattern sheet\n'
                    else:
                        logging.info('Info :: Signal Data sheet not found in Test pattern sheet')
                        error_Count = error_Count + 1
                        Missing_Input_Details += str(error_Count) + '. Signal Data sheet not found in Test pattern sheet\n'
                else:
                    print "Test pattern sheet not found"
                    logging.info('Info :: Test pattern sheet not found ')
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + '. Test pattern sheet not found\n'
            else:                                                                                                               # if vehicle or keywords not found it will display message
                print "Vehicle not found master test sheet folder"
                logging.info('Info :: Vehicle not found master test sheet folder ')
                error_Count = error_Count + 1
                Missing_Input_Details += str(error_Count) + '. Vehicle folder not found in master test sheet folder\n'
                #tkMessageBox.Showinfo("Info","Vehicle not found in Master Test Sheet")
        except:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt',
            level=logging.INFO,format='folder not present')                        
            pass




        
        Test_Sheet = Book_Master_TP        
        print "Test_Sheet",Test_Sheet
        Book_Master_Judgement_sheet = Org_Path + "06_Master_Judgement_Sheet" + "\\" + "Master_Judgement_Sheet_FLS.xls"
        book = xlrd.open_workbook(Test_Sheet)                                                                       # Opening the test case sheet
        sig_data_sheet = book.sheet_by_name(sig_data_sheet_str)


        Vehicle_Name = VehicleName + ' - ' + RegionName + ' - ' + PartNo       

        xlapp1 = win32com.client.Dispatch("Excel.Application")   #To open Excel 

        if os.path.exists(str(interface_sheet_path)):

            xlapp1.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    
            xlapp1.Application.Run("Interface_VBA.xls!module7.Change_DBC_Name")

        xlapp1.Workbooks.Close()
        

        curItem=0
       
        try:

            #VehicleDict[VehicleName][RegionName][PartNo]= OrderedDict()
            #Covariant=[]
            count=0
            Enabled_Testcase_name =[]
            Variant=[]
            ApplicationEnabled=[]
            TestEnabled=[]
            CoArray=[]
            Mode_Value = []
            #Variant_Value= []
            TestCaseEnabled=[]
            var=[]
            CoRow = 0
            CoCol= 0
            DIMPSheet=WorkBook.sheet_by_name("ICC_Cancel_Testing")                                     
            DIMPSheetCol = DIMPSheet.ncols
            DIMPSheetRow = DIMPSheet.nrows
            print "DIMPSheetRow",DIMPSheetRow
            for i in range (0, DIMPSheetRow ):                                                                               # Loop for finding last used row in ICC_Cancel_Testing sheet
                for j in range(0,DIMPSheetCol):
                    if DIMPSheet.cell(i,j).value!='':
                        xfx=DIMPSheet.cell_xf_index(i,j)                        
                        xf=WorkBook.xf_list[xfx]
                        pattern=xf.background.pattern_colour_index
                        background=xf.background.background_colour_index
                        if pattern==13 and background==64:
                            if DIMPSheet.cell(i,j).value=='ICC END':
                                ItsEndRow = i
                                break;
                                
            print "ItsEndRow",ItsEndRow        
            for i in range (0, ItsEndRow ):                                                                               # Loop for finding Variant code column
                for j in range(0,DIMPSheetCol):
                    if DIMPSheet.cell(i,j).value!='':
                        xfx=DIMPSheet.cell_xf_index(i,j)
                        
                        xf=WorkBook.xf_list[xfx]
                        pattern=xf.background.pattern_colour_index
                        background=xf.background.background_colour_index
                        if pattern==13 and background==64:
                            if DIMPSheet.cell(i,j).value=='Variant Code':
                                
                                CoRow =  i
                                CoCol= j

            TestCaseIdCol =  CoCol                              
            print "CoRow",CoRow
            for j in range(CoCol + 1,DIMPSheetCol):                                                                 # Loop for storing Variant code column
                    if DIMPSheet.cell(CoRow,j).value!='':
                        xfx=DIMPSheet.cell_xf_index(CoRow,j)
                        xf=WorkBook.xf_list[xfx]
                        pattern=xf.background.pattern_colour_index
                        background=xf.background.background_colour_index
                        if pattern==13 and background==64:
                            #Covariant.append(DIMPSheet.cell(CoRow,j).value)                                         
                            #Variant_Value.append(DIMPSheet.cell(CoRow+1,j).value)
                            Mode_Value.append(DIMPSheet.cell(CoRow+2,j).value)
                            CoArray.append(j)                                                                       # This array stores the Variant code column for Vehicle 
        

            for i in range(0,len(Covariant)):                                                                       
                Variant.append(PartNo+ '_' + str(int(Variant_Value[i])))                                            # Appending the Variant_Value to part number                                                                                                                                      # Variant array contains the variant name (Variant_Value and part number)

            #for variants in Variant:
            #    VehicleDict[VehicleName][RegionName][PartNo][variants]=OrderedDict()
##            print "Covariant",Covariant
##            print "Variant_Value",Variant_Value

            Vehicle_Details = VehicleName +  '_' + RegionName
            Vehicle_Name = VehicleName + '_' + RegionName + '_' + PartNo
            Actual_PartNo=PartNo.split("_")[1]




            ICC_CANCEL_Dict = OrderedDict()                    # This makes dictionary required for ICC_CANCEL Tree
            ICC_CANCEL_Dict[Vehicle_Details]= OrderedDict()
            ICC_CANCEL_Dict[Vehicle_Details][PartNo]= OrderedDict()
            print "ICC_CANCEL_Dict",ICC_CANCEL_Dict
            Active_Test = 'ICC_Cancel_Testing'       
            ICC_CANCEL_Dict[Vehicle_Details][PartNo][Active_Test]=OrderedDict()
            Failsafe_Result_Folder=[]

            TestCaseEnabled = []
            Enabled_Testcase = ''

            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
            logging.info('Dispatch Sheet for Failsafe ADAS_2.5 Loaded')


           # print "len(Variant)",len(Variant)
            for k in range(0,len(Variant)):                                                                     #This loop for number of variant presents.

                print "k",k,int(Variant_Value[k])
                VariantName_tree ='Variant_' +  str(Variant_Value[k])
                ICC_CANCEL_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree]= OrderedDict()                   #Store Variant in Dictionary
                Mode_value = DIMPSheet.cell(CoRow + 2,TestCaseIdCol+k+1).value
                #print "Mode_value",Mode_value
                ICC_CANCEL_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Mode_value]= OrderedDict()      #Store Mode in Dictionary
                for i in range (CoRow + 3, ItsEndRow):      # ( Var_Row +2 ) contains the string "Y" or "N"                   
                    xfx=DIMPSheet.cell_xf_index(i,CoArray[k])
                    xf=WorkBook.xf_list[xfx]
                    pattern=xf.background.pattern_colour_index
                    background=xf.background.background_colour_index
                    if pattern==9 and background==64:
                        TestCaseEnabled.append(DIMPSheet.cell(i,TestCaseIdCol).value)
                        Enabled_Testcase =DIMPSheet.cell(i,TestCaseIdCol).value
                        Enabled_Testcase_name.append(DIMPSheet.cell(i,TestCaseIdCol-1).value)
                       #print "Enabled_Testcase",Enabled_Testcase

                    
                        #print "ICC_CANCEL_Dict1",ICC_CANCEL_Dict
                        ICC_CANCEL_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Mode_value][Enabled_Testcase]=OrderedDict()      #Store Enable Test case name in Dictionary
                        ##print "ICC_CANCEL_Dict2",ICC_CANCEL_Dict
                   


    

            uid = 0  #added
            ICC_Cancel_Tree = construct_JSON_tree(ICC_CANCEL_Dict,frame12)                      #Make Tree from genrated Dictionary
            uid_MSG_prev=0
            curItem = 0
            Var_Val=0
            ManeuverTesting_vehicle_id_entry["state"] = NORMAL
            ManeuverTesting_vehicle_id_entry.delete(0, END)
            ManeuverTesting_vehicle_id_entry.insert(0, Vehicle_Details)
            ManeuverTesting_vehicle_id_entry["state"] = DISABLED
            def INPUT_SIGNAL_replace(Failsafe_Application_sheet,ApplicationSheet_TestCase_Start_Row,Failsafe_Application_sheet_End_col):           #This function is to Extract Input signal data from master test pattern
                global CAN_Value_1,CAN_Value_2
                INPUT_SIGNAL_1_str= "INPUT_SIGNAL_1"

                INPUT_SIGNAL_1_column	= 10
                INPUT_SIGNAL_2_column   = 11
               # print "input",ApplicationSheet_TestCase_Start_Row,Failsafe_Application_sheet_End_col
                try:
                    for i in range (0,  Failsafe_Application_sheet_End_col):             # Loop for traversing through the EXCEL sheet
                        #print "data",Failsafe_Application_sheet.cell(0,i).value
                        if Failsafe_Application_sheet.cell(0,i).value == INPUT_SIGNAL_1_str:         #This finds row and column of the INPUT_SIGNALS
                            INPUT_SIGNAL_1_column =  i
                            INPUT_SIGNAL_2_column = INPUT_SIGNAL_1_column+1
                            print "INPUT_SIGNAL_1_column",INPUT_SIGNAL_1_column,INPUT_SIGNAL_2_column                           
                except:
                    INPUT_SIGNAL_1_column=10

                    INPUT_SIGNAL_2_column=11


                try:
                    CAN_Value_1=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_1_column).value  
                except:
                    CAN_Value_1="NA"
                try:
                    CAN_Value_2=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_2_column).value
                except:
                    CAN_Value_2="NA"
                print "CAN_Value_1,CAN_Value_2",CAN_Value_1,CAN_Value_2
                return  CAN_Value_1,CAN_Value_2                           

            def CAN_Message_Replacement(CAN_Value_1,CAN_Value_2):                                 #This function is to write Input signal data from master test pattern to Py-mscript of judgement sheet
                book1 = xlrd.open_workbook(str(Book_Master_Judgement_sheet),formatting_info=True)
                sheetNames = book1.sheet_names()
                sheetNumber = 0
                for i in sheetNames:                                                                #This loop to Extract sheet number
                    if "Py-MScript" in i:
                        break
                    sheetNumber = sheetNumber + 1
                        
                book2 = copy(book1)                                                                 # Write CAN_Value_1 and CAN_Value_2 to judgement sheets Py-mscript
                Interface_sheet = book2.get_sheet(sheetNumber)                                      
                Interface_sheet.write(25, 2, CAN_Value_1)
                Interface_sheet.write(25, 3, CAN_Value_2)
           
                book2.save(Book_Master_Judgement_sheet)
                
            def Make_SignalData_Array(Procedure_Test_sheet,x):                               #This function is to create SigInfo_temp,SignalData1_temp list for passing it to Execution function
                global SignalData1_temp,SigInfo_temp,sig_info
                SigNames=[]
                sig_name=None
                sig_delay=None
                sig_path=None
                sig_value=None
                sig_reset=None
                
                sig_name = Procedure_Test_sheet.cell(x,0).value
                sig_delay = Procedure_Test_sheet.cell(x,2).value   
                sig_path = Procedure_Test_sheet.cell(x, 6).value
                sig_value = Procedure_Test_sheet.cell(x,7).value
                sig_reset = Procedure_Test_sheet.cell(x,8).value
                SigNames.append(sig_name)    
                sig_data = [sig_path,sig_value,sig_reset]                  
                sig_info[sig_name] = sig_data
                SigInfo_temp.append(sig_info[sig_name])
                SignalData1_temp =data_function(Procedure_Test_sheet,sig_path,sig_value,sig_name,sig_delay,sig_reset,x)                    # pass all extracted value to data_function
                return SigInfo_temp,SignalData1_temp

            def data_function(Primary_Test_sheet,sig_path,sig_value,sig_name,sig_delay,sig_reset,x):     #This function get signal related data from test pattern with respect to exec cont. 
                signal_data = []
                save_var = None
                if str(Primary_Test_sheet.cell(x,1).value) in \
                   execute_start_end_str:                                                                       # If string in cloumn 2 is 'exec_start_end' then call 'get_data_execute_start_end_str' to collect test case data 
                    signal_data = get_data_execute_start_end_str(sig_name,
                                                                 sig_path,
                                                                 Primary_Test_sheet,
                                                                 sig_delay,x,
                                                                 sig_data_sheet)
                    save_var = 1
                elif str(Primary_Test_sheet.cell(x,1).value) in execute_cont_str:                                # If string in cloumn 2 is 'exec_cont' then call 'get_data_execute_conti' to collect test case data
                                  
                    signal_data = get_data_execute_cont(sig_name,sig_path,
                                                        Primary_Test_sheet,
                                                        sig_delay,x,
                                                        sig_data_sheet)
                    save_var = 1

                
                   
                elif str(Primary_Test_sheet.cell(x,1).value) in exec_var_dep:                                    # If string in cloumn 2 is 'exec_var_dep' then call 'get_data_execute_var_dep' to collect test case data   
                    signal_data = get_data_execute_var_dep(sig_name,sig_path,
                                                           Primary_Test_sheet,
                                                           sig_delay,x,
                                                           sig_data_sheet)
                    save_var = 1
                    
                elif str(Primary_Test_sheet.cell(x,1).value) in exec_delay:                                      # If string in cloumn 2 is 'exec_delay' then call 'get_execute_delay' to collect test case data  
                    signal_data = get_execute_delay(sig_name,sig_delay)
                    save_var = 1


                SignalData.append(signal_data) 
            


                if save_var == 1:                                                                               # Save_var = 1 signifies that atleast 1 signal is present in that test case 
                    save_var = 0
                    
                signal_data = []
                sig_info = {}
                return SignalData
            #************ Screenshot function**************#
# Brings forward Control Desk application. Takes screenshot of Meter Navi.layout and Diag layout.#			

            def Screenshot(screenshot_path,Meter_Navi_Layout_Name,Diag_layout_Name,Sub_TestCaseName_tree):      # This function is to take screenshot of DTC and meter navi
                global ICC_Cancel_code
                time.sleep(1)
                try:      
                    wildcard = ".*ControlDesk Developer Version*"
                    cW = cWindow()
                    handle_manager=cW.find_window_wildcard(wildcard)
                    cW.Maximize()
                    cW.BringToTop()
                    cW.SetAsForegroundWindow()
                    bbox = win32gui.GetWindowRect(handle_manager)

                except:
                    f = open("log.txt", "w")
                    f.write(traceback.format_exc())
                    print traceback.format_exc()

                time.sleep(1)                
                try:
                    print "FilePath2",FilePath2,Meter_Navi_Layout_Name
                    Layout = FilePath2 + "\\" +Meter_Navi_Layout_Name
                    Instrumentation().Layouts.Item(Layout).Activate()  
                    Instrumentation().ActiveLayout.Maximize()
                    snapshot=ImageGrab.grab(bbox)            
                    snapshot.save(screenshot_path+"\\"+Sub_TestCaseName_tree+"_Meter_navi.jpg")
                    time.sleep(2)
                    Layout = FilePath2 + "\\" + Diag_Layout_Name
                    Instrumentation().Layouts.Item(Layout).Activate()  
                    Instrumentation().ActiveLayout.Maximize()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(8)                                              
                    time.sleep(1.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    #for dtc_array_count in range(0,9):
                    print " DTC_string_path_temp_ICC",DTC_string_path_temp_ICC
                    Actual_DTC_set = myAppl.Variable(str(DTC_string_path_temp_ICC)).Read()                              #read cancel code
                    print "Actual_DTC_set",Actual_DTC_set,ICC_Cancel_code
                    Actual_DTC_set =hex(int(Actual_DTC_set))
                    print "Actual_DTC_set",Actual_DTC_set,ICC_Cancel_code
                    #Actual_DTC_set[dtc_array_count] = myAppl.Variable(str(DTC_string_path_temp + str(dtc_array_count + 1) + "{SubArray1}")).Read()
                    #Actual_DTC_subArray_set[dtc_array_count] = myAppl.Variable(str(DTC_string_path_temp + str(dtc_array_count + 1) + "{SubArray2}")).Read()
                    if str(Actual_DTC_set) == str(ICC_Cancel_code) or ICC_Cancel_code == "NA" :                                                        #compare read code and given code
                        ICC_Cancel_read = 1
                    else:
                        ICC_Cancel_read = 0

                    print "ICC_Cancel_read",ICC_Cancel_read                        
                    book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                    sheetNames = book1.sheet_names()
                    sheetNumber = 0
                    for i in sheetNames:
                        if "Py-MScript" in i:
                            break
                        sheetNumber = sheetNumber + 1
    #                    print "SheetNumber",sheetNumber
                    book2 = copy(book1)
                    Interface_sheet = book2.get_sheet(sheetNumber)
                    print "sheetNumber",sheetNumber
                    Interface_sheet.write(7, 0, 'ICC_Cancel_read')                                                      
                    Interface_sheet.write(7, 1, ICC_Cancel_read)                                                            #write ICC_Cancel_read value to master judgement sheet  
                    book2.save(Book_Master_Judgement_sheet)                        
                                            
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)
                    if ICC_Cancel_read == 1:
                        snapshot=ImageGrab.grab(bbox)                                                                       #Taking screenshot only if ICC code is correct
                        snapshot.save(screenshot_path+"\\"+Sub_TestCaseName_tree+"_DTC.jpg")
                    time.sleep(2)

                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('Meter_Navi_Layout_Name,Diag_Layout_Name or DIAG_CMD_NO_path not found')




                try:      
                    wildcard = ".*CANape"
                    cW = cWindow()
                    cW.find_window_wildcard(wildcard)
                    cW.Maximize()
                    cW.BringToTop()
                    cW.SetAsForegroundWindow()
                    time.sleep(1)

                except:
                    f = open("log.txt", "w")
                    f.write(traceback.format_exc())
                   ## print traceback.format_exc()


            def Reset_functionality(TestCase_Start_Row,TestCase_End_Row,Primary_Test_sheet,ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row,Failsafe_Application_sheet,myAppl):      # This function is for resetting 
                try:
                    for m in range(ApplicationSheet_TestCase_Start_Row+1,ApplicationSheet_TestCase_End_Row ):            #extract resetting data from appication sheet of master test sheet      
                        if Failsafe_Application_sheet.cell(m,6).value!='' :
                           
                            set_sig_reset_value = None
                            set_sig_path = None
                            set_sig_default_value = None
                            set_sig_Appl = None
                            set_sig_reset_value = int(Failsafe_Application_sheet.cell(m,8).value)                  
                            set_sig_path = Failsafe_Application_sheet.cell(m,6).value
                            set_sig_default_value = Failsafe_Application_sheet.cell(m,7).value                   
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(0.5)
                            

                    for m in range(TestCase_Start_Row + 1,TestCase_End_Row):                                    #extract resetting data from primary sheet of master test sheet 
                       
                        if Primary_Test_sheet.cell(m,6).value != '' :
                           
                            set_sig_reset_value = None
                            set_sig_path = None
                            set_sig_default_value = None
                            set_sig_Appl = None
                            set_sig_reset_value = int(Primary_Test_sheet.cell(m,8).value)
                            set_sig_path = Primary_Test_sheet.cell(m,6).value
                            set_sig_default_value = Primary_Test_sheet.cell(m,7).value
                            myAppl.Variable(set_sig_path).Write(0)
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)                            
                            time.sleep(0.5)
                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('signal Path , reset value not found in sheet')
             
                try:   

                    # clearing DTC values
                    myAppl.Variable(Power_Supply_path).Write(129)
                    time.sleep(2)
                    myAppl.Variable(Power_Supply_path).Write(1)
                    time.sleep(2)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(3)           # clear DTC
                    time.sleep(1.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)
                    
                    myAppl.Variable(DIAG_CMD_NO_path).Write(8)
                    time.sleep(1)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)
                    
                    myAppl.Variable(Power_Supply_path).Write(129)
                    time.sleep(2)
                    myAppl.Variable(Power_Supply_path).Write(1)
                    time.sleep(2)

                    myAppl.Variable(DIAG_CMD_NO_path).Write(2)          # read DTC
                    time.sleep(0.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)


                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('power supply & Diag Paths are wrong....... Check signal paths in signal data sheet')            



            ICC_Cancel_overall_progressbar["maximum"]=uid           
            ICC_Cancel_overall_progressbar["value"] = 0 
            for j in range(0, uid):                                 # iteration for number items present in tree
                overall_progressbar_uid=j
                ICC_Cancel_overall_progressbar["value"] = overall_progressbar_uid    
                ICC_Cancel_overall_progressbar["value"] = j	
                Ances_array=[]
                TestCaseNameId = ''
                application_name = ''
                ApplicationNameId = ''
                TestCaseNameId = ''
                VariantName_tree = ''
                ApplicationName_tree = ''
                ManeuverTesting_result_entry.delete(0, END)
                TestCaseName_tree = ''
                ManeuverTesting_testcase_entry.delete(0, END)
                ManeuverTesting_application_entry.delete(0, END)
                ManeuverTesting_variant_entry.delete(0, END)
                curItem= curItem + 1
                ICC_Cancel_Tree.selection_set(curItem)
                Item=curItem
                Master_CANape_Path = Org_Path + '04_Master_CANape_Configurations\\05_Failsafe_CANape_Configuration'    # Creating path for source folder to copy CANape files from Master folder
                book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                sheetNames = book1.sheet_names()
                sheetNumber = 0

                for i in range(0, 10):
                    src_folder = Master_CANape_Path
                    
                    dest_folder = folder_path                                                                   # Path for destination folder to copy CANape files from Master folder
                    
                    ParentItem=ICC_Cancel_Tree.parent(Item)
                    Ances_array.append(ParentItem)
                    Item=Ances_array[i]
                  
                    if Ances_array[i]=='':
                        break
                    else :
                        continue
                
                #for id1 in Ances_array:                                                                            # Functionality to open the tree
                  #  tree.item(id1, open=True)       
                #print VariantName_tree   
                Heir = len(Ances_array)
                #print Heir,"Heir"
                #print Ances_array,"Ances_array"
                #ManeuverTesting_overall_progressbar["maximum"] =  Heir                                                                      # Progressbar max length assigned
                if Heir == 6:
                    

                    TestCaseNameId = curItem
                    ApplicationNameId = Ances_array[0]
                    VariantNameId = Ances_array[1]
                    
                elif Heir == 5:
                   
                    ApplicationNameId = curItem
                    VariantNameId = Ances_array[0]
                    TestCaseNameId = ''
                    
                elif Heir == 4:
                    VariantNameId = curItem
                    ApplicationNameId = ''
                    TestCaseNameId = ''
                elif Heir == 3:
                    print "curItem",curItem
                    application_name = curItem
                    VariantNameId = ''
                    ApplicationNameId = ''
                    TestCaseNameId = ''
                else:
                    VariantNameId = ''
                    ApplicationNameId = ''
                   
                    TestCaseNameId = ''
                
                if VariantNameId=='':
                    ITS_variant_entry.delete(0, END)
                else:
                    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet 

                    if os.path.exists(str(interface_sheet_path)):
                        xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)                    
                        xlapp.Application.Run("Interface_VBA.xls!module14.Copysheet",'ICC_CANCEL')
                    xlapp.Workbooks.Close()                    
                    
                    VariantName_tree = ICC_Cancel_Tree.item(VariantNameId, 'text')
                    print VariantNameId,"VariantNameId"
                    print "VariantName_tree",VariantName_tree

##                        VariantName_tree =  str( VariantName_tree)
##                      for i in range(0,len(Covariant)):
                        ##VariantName_tree = Variant_Value[i]
                    data  = VariantName_tree.split('_')
##                        VariantName_tree = VariantName_tree[-3:]
##                        if '_' in VariantName_tree:
##                            VariantName_tree = VariantName_tree[-2:]

                    VariantName_tree = ' Variant ' + data[1]
                    ManeuverTesting_variant_entry["state"] = NORMAL
                    ManeuverTesting_variant_entry.delete(0, END)
                    ManeuverTesting_variant_entry.insert(0, VariantName_tree)
                    print 'present variant is ' + VariantName_tree
                    ManeuverTesting_variant_entry["state"] = DISABLED
                    
                if ApplicationNameId=='':
                    ManeuverTesting_application_entry.delete(0, END)
                else:
                    ApplicationName_tree = ICC_Cancel_Tree.item(ApplicationNameId, 'text')
                    ManeuverTesting_application_entry["state"] = NORMAL
                    ManeuverTesting_application_entry.delete(0, END)
                    print 'present appli is ' + ApplicationName_tree
                    ManeuverTesting_application_entry.insert(0, ApplicationName_tree)
                    ManeuverTesting_application_entry["state"] = DISABLED

                if application_name=='':
                    ManeuverTesting_application_entry.delete(0, END)
                else:                                                                                   
                    application_name_tree = ICC_Cancel_Tree.item(application_name, 'text')
                    ManeuverTesting_application_entry["state"] = NORMAL
                    ManeuverTesting_application_entry.delete(0, END)
                    print 'application_name_tree ' + application_name_tree
                    ManeuverTesting_application_entry.insert(0, application_name_tree)
                    ManeuverTesting_application_entry["state"] = DISABLED
                    
                if TestCaseNameId == '':
                    ManeuverTesting_testcase_entry.delete(0, END)
                else:
                    TestCaseName_tree = ICC_Cancel_Tree.item(TestCaseNameId, 'text')
                    ManeuverTesting_testcase_entry["state"] = NORMAL
                    ManeuverTesting_testcase_entry.delete(0, END)
                    ManeuverTesting_testcase_entry.insert(0, TestCaseName_tree)
                    print 'present test case is ' + TestCaseName_tree
                    ManeuverTesting_testcase_entry["state"] = DISABLED
                    Log_File_Path = Script_Path +'\\HILS_Testing_Log.txt'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Test case is - %s',TestCaseName_tree)
                        
                if TestCaseName_tree == '' and ApplicationName_tree == '' and \
                   VariantName_tree == '':
                    print ' While Vehicle features '   
                    time.sleep(1)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Vehicle is - %s',Vehicle_Id)
                    
                    continue
               
                elif TestCaseName_tree == ''  and ApplicationName_tree == ''and\
                     VariantName_tree!= '' :
                    print 'Var_Val', Var_Val
##                        Instrumentation().Layouts.Item(LayoutConfig).Activate()
##                        Instrumentation().ActiveLayout.Maximize() 
                    myAppl.Variable(Power_Supply_path).Write(1)
                    Write_Var = Variant_Value[Var_Val]
                    print " Write_Var: ", Write_Var
                    ##Instrumentation().ActiveLayout.Normalize()
                    Variant_write(myAppl,Write_Var)                                         

                    Var_Val = Var_Val + 1
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Variant is - %s', VariantName_tree) 
                    time.sleep(3)
                        
                    continue
                
                elif TestCaseName_tree == ''  and ApplicationName_tree != '' and VariantName_tree!= '':
                   
                    ApplicationName = application_name_tree.rsplit('_',1)[0]
                    print ApplicationName,"ApplicationName"
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Application is - %s', ApplicationName_tree)

                    Sub_VariantName_tree = str(int(float(data[1])))
                    print "Sub_VariantName_tree",Sub_VariantName_tree
                    Report_src_folder = Org_Path + '05_Master_Result_Reports' + "\\" +  ApplicationName
                    
                    print 'Report_src_folder' + Report_src_folder
                    print  'dest_folder' + dest_folder
                    
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    print dird,"dird"
                    if not '_' in Sub_VariantName_tree :
                        
                        Sub_VariantName_tree = '_' + Sub_VariantName_tree
                   
                    print "Sub_VariantName_tree",Sub_VariantName_tree


                    
                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break

                    
                    #folder_name_app = ApplicationName
                    dest_folder = dest_folder +'\\' +  folder_name_app


                    Report_dest_folder = dest_folder
                    print 'Report_dest_folder ' + Report_dest_folder
                    distutils.dir_util.copy_tree(Report_src_folder,Report_dest_folder)
                    rep_name = ApplicationName + '.xls'
                    Rep_Name = ApplicationName  + Sub_VariantName_tree+ '.xls'
                    os.rename(os.path.join(Report_dest_folder,rep_name),os.path.join(Report_dest_folder,Rep_Name))
                    Rep_Path = os.path.join(Report_dest_folder,Rep_Name)
                    print Rep_Path
##                    interface_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Interface_VBA.xls'
                    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet 

                    if os.path.exists(str(interface_sheet_path)):

                        xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    
                        xlapp.Application.Run("Interface_VBA.xls!module5.Hide_All",Rep_Path)


                    xlapp.Workbooks.Close()
                    Application_sheet = book.sheet_by_name(ApplicationName)
                    End_row = Application_sheet.nrows
                    Data_End_Row =  sig_data_sheet.nrows
##                        Act_lay = Application_sheet.cell(0,1).value
                    


                    time.sleep(1)

                    continue
                
                else:
                    
                    print "Enabled_Testcase_name[i]",Enabled_Testcase_name[count]
                    
                    
                    ICC_Cancel_progressbar["maximum"] = 6
                    ICC_Cancel_progress=1
                    ICC_Cancel_progressbar["value"] = ICC_Cancel_progress  
                    flag = 0
                    Sub_TestCaseName_tree = TestCaseName_tree[-2:]
                    ICC_Cancel_Tree.item(curItem, text = ApplicationName + '_TEST_' + Sub_TestCaseName_tree, values = Enabled_Testcase_name[count])
                    #ICC_Cancel_Tree.item(curItem, text = ApplicationName + '_TEST_' + Sub_TestCaseName_tree,values = (Enabled_Testcase_name[count],"Result_Strng") )
                    ssdest_folder = ''
                    #print "Sub_TestCaseName_tree",Sub_TestCaseName_tree
                    #print 'Src_Folder' + src_folder


                    print "ApplicationName_tree",ApplicationName_tree

                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
####                        Sub_VariantName_tree = '_' + Sub_VariantName_tree
                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break
                    dest_folder = dest_folder +'\\' +  folder_name_app
                   # print "dest_folder",dest_folder
##                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
##                    for directories_d in dird:
##                        if ApplicationName_tree == directories_d[3:]:
##                            dest_folder = dest_folder + '\\' + directories_d
##                            print "ApplicationName_tree dest_folder",dest_folder
##                            ssdest_folder = dest_folder
##                            break
                    dest_folder = dest_folder + '\\' + TestCaseName_tree
                   # print "dest_folder",dest_folder
                    
                    ssdest_folder = dest_folder
                    book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                    sheetNames = book1.sheet_names()
                    sheetNumber = 0
                    for i in sheetNames:
                        if "Py-MScript" in i:
                            break
                        sheetNumber = sheetNumber + 1
    #                    print "SheetNumber",sheetNumber
                    book2 = copy(book1)
                    Interface_sheet = book2.get_sheet(sheetNumber)
                    Interface_sheet.write(0, 1, VehicleName)
                    print "VehicleName",VehicleName
                    Interface_sheet.write(1, 1, RegionName)
                    print "RegionName",RegionName
                    Interface_sheet.write(2, 1, Write_Var)
                    Interface_sheet.write(3, 1, ApplicationName)  ##ApplicationName_tree
                    print "ApplicationName",ApplicationName

                    Interface_sheet.write(4, 1, Sub_TestCaseName_tree)
                    Interface_sheet.write(8, 1, dest_folder)
                    Interface_sheet.write(30, 0, 'Meter_Navi_Enabled_Applications')
                    Interface_sheet.write(31, 0, 'Meter_Navi_Enabled_Applications_End')

                    book2.save(Book_Master_Judgement_sheet)

                    ITS_CANape_src_folder = src_folder +'\\01_Failsafe_reference_folder'

                    try:
                        os.mkdir(dest_folder)
                    except:
                        pass

                    screenshot_folder=dest_folder +'\\00_Screenshot'
                    print dest_folder
                    try:
                        os.mkdir(screenshot_folder)
                    except:
                        pass
                    
                    if (flag == 1):
                        print "flag in if ",flag
                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.info("Test case not found in Master Configuration %s" + TestCaseName_tree + "  %s" + Sub_TestCaseName_tree)
                    else:
                     
                        distutils.dir_util.copy_tree(ITS_CANape_src_folder,dest_folder)

                    
                    Application_sheet = book.sheet_by_name("Failsafe_Primary")

                End_row = Application_sheet.nrows
                Data_End_Row =  sig_data_sheet.nrows
                end_test_str = 'end_set_ICC_test_case'
                start_test_str = 'start_set_ICC_test_case'
                print end_test_str,start_test_str      
                for k in range(0, End_row):
                    
                    if Application_sheet.cell(k,0).value== start_test_str :
                        TestCase_Start_Row= k
                        break
                         
                    else:
                        continue
                for k in range(0, End_row):
                    if Application_sheet.cell(k,0).value== end_test_str :
                        TestCase_End_Row= k
                        break
                    else:
                        continue

                for k in range(0, End_row):
                    if Application_sheet.cell(k,0).value== Meter_Navi_Layout_str :
                        Meter_Navi_Layout_Name= Application_sheet.cell(k,1).value
                        break

                for k in range(0, End_row):
                    if Application_sheet.cell(k,0).value== Diag_layout_str :
                        Diag_Layout_Name= Application_sheet.cell(k,1).value
                        break
                                
                ICC_Cancel_Application_sheet = book.sheet_by_name(ApplicationName)
                ICC_Cancel_Application_sheet_End_row = ICC_Cancel_Application_sheet.nrows
                ICC_Cancel_Application_sheet_End_col = ICC_Cancel_Application_sheet.ncols
                Data_End_Row =  ICC_Cancel_Application_sheet.nrows
                TestS = TestCaseName_tree[-2:]
#                print "TestS",TestS
                start_test_str = 'start_'+ ApplicationName +'_Test_' + TestS
                end_test_str = 'end_'+ ApplicationName +'_Test_' + TestS
                print end_test_str,start_test_str
                
                for k in range(0, ICC_Cancel_Application_sheet_End_row):
                    
                    if ICC_Cancel_Application_sheet.cell(k,0).value== start_test_str :
                        ApplicationSheet_TestCase_Start_Row= k
                        break
                         
                    else:
                        continue
                for k in range(0, ICC_Cancel_Application_sheet_End_row):
                    if ICC_Cancel_Application_sheet.cell(k,0).value== end_test_str :
                        ApplicationSheet_TestCase_End_Row= k
                        break
                    else:
                        continue
                ICC_Cancel_progress=ICC_Cancel_progress+1
                ICC_Cancel_progressbar["value"] = ICC_Cancel_progress
                
                INPUT_SIGNAL_replace(ICC_Cancel_Application_sheet,ApplicationSheet_TestCase_Start_Row,ICC_Cancel_Application_sheet_End_col)
                print "CAN_Value_1,CAN_Value_2",CAN_Value_1,CAN_Value_2
                CAN_Message_Replacement(CAN_Value_1,CAN_Value_2)
                
                print TestCase_Start_Row,TestCase_End_Row
                save_var = 0
                Primary_Test_sheet = Application_sheet
                for x in range(TestCase_Start_Row + 1,TestCase_End_Row):
                    SigInfo,SignalData1=Make_SignalData_Array(Primary_Test_sheet,x)

 #               print "SignalData",SignalData                    
#                print "ApplicationSheet_TestCase_Start_Row",ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row
                Procedure_Test_sheet = ICC_Cancel_Application_sheet
                for x in range(ApplicationSheet_TestCase_Start_Row+1,ApplicationSheet_TestCase_End_Row):                                
                    Procedure_Test_sheet=ICC_Cancel_Application_sheet                                                            
                    SigInfo,SignalData1=Make_SignalData_Array(Procedure_Test_sheet,x)                    

#                print "SignalData",SignalData

##                    interface_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Interface_VBA.xls'
                xlapp = win32com.client.Dispatch("Excel.Application")   #To open Excel 

                if os.path.exists(str(interface_sheet_path)):

                    xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)

                    xlapp1.Application.Run("Interface_VBA.xls!module13.replace_data_WithoutCAN")
                    
                    xlapp.Application.Run("Interface_VBA.xls!module11.Make_Canape_failsafe")
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(message)s')
                logging.info('CANape Log copy paste done')

                xlapp.Workbooks.Close()

                book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)                                   
                JudgementsheetName = book1.sheet_by_name("ICC_CANCEL")
                if JudgementsheetName.cell(9,1).value == "NA":                                                      #check in judgementsheet for "A" or  "NA" for testcase applicable. if testcase not applicable it will disply TBD on GUI
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Test case skipped due to test case is not Applicable in judgementsheet %s')
                    ICC_Cancel_Tree.item(curItem, text = ApplicationName_tree + '_TEST_' + Sub_TestCaseName_tree, values = "TBD" )
                    continue 

                book5 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                ICC_Application_sheet = book5.sheet_by_name("ICC_CANCEL")
                ICC_Application_sheet_end_row = ICC_Application_sheet.nrows
                ICC_Application_sheet_end_col = ICC_Application_sheet.ncols
                start_Judge_test_str = 'start_keyword_' + TestS
                for row in range(0,ICC_Application_sheet_end_row):
                    if ICC_Application_sheet.cell(row,0).value ==  start_Judge_test_str:
                        ICC_Cancel_code = ICC_Application_sheet.cell(row,14).value
                        print "ICC_Cancel_code",ICC_Cancel_code
                        break;
                
                
                write_ss_info = open(dest_folder + "\\Screen_Shot.txt","w")
                write_ss_info.write(ssdest_folder + "\\" + "00_Screenshot\n")
                write_ss_info.write(TestCaseName_tree[-2:])
                write_ss_info.write("_")
                write_ss_info.write(ApplicationName)  ##ApplicationName_tree
                write_ss_info.close()

            

                
            

                pathTextFile = dest_folder
                pathTextFile = pathTextFile + "\\" + "Sync.txt"
                print "pathTextFile", pathTextFile


                myAppl.Variable(Power_Supply_path).Write(1)
                time.sleep(1)
                myAppl.Variable("simState").Write(0)                                                                    # 'Reset' Simstate
                time.sleep(.5)
                myAppl.Variable("simState").Write(2)                                                                    # 'Set' Simstate
                time.sleep(2)

                myAppl.Variable(Power_Supply_path).Write(129)
                time.sleep(1)
                myAppl.Variable(Power_Supply_path).Write(1)                                                                
                time.sleep(6)
                
                myAppl.Variable(DIAG_CMD_NO_path).Write(3)                                                          #Clear DTC
                time.sleep(.5)                  
                myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                time.sleep(.5)
         
                flag_speed = 0
                count = 0
                while(flag_speed == 0):                                                                             #check for vehicle speed and DTC 
                    read_velocity_1 = myAppl.Variable(Read_vehicle_speed_path).Read()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(2)                                                          #Read DTC
                    time.sleep(0.5)                 
                    Actual_DTC_set = myAppl.Variable(DTC_string_path_temp).Read()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(.5)
                    print "Actual_DTC_set",int(float(Actual_DTC_set))
                    time.sleep(7)
                    read_velocity_2 = myAppl.Variable(Read_vehicle_speed_path).Read()
                    print "vehicle speed read",int(float(read_velocity_2)),int(float(read_velocity_1))
                    if (int(float(read_velocity_2)) - int(float(read_velocity_1)) > 0) and (int(float(Actual_DTC_set)) == 0) :                    
                        flag_speed = 1
                    else:
                        myAppl.Variable(Power_Supply_path).Write(129)
                        time.sleep(1)
                        myAppl.Variable(Power_Supply_path).Write(1)                                                                 # 'Set' Simstate
                        time.sleep(6)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(3)
                        time.sleep(.5)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                        time.sleep(.5)                        
                        myAppl.Variable(DIAG_CMD_NO_path).Write(2)
                        time.sleep(0.5)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                        time.sleep(0.5)

                    if count > 5:    
                        flag_speed = 1
                    count = count + 1

                if count > 5:
                    Controldesk_Load_Reload(Write_Var)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('DTC not cleared')
                    #continue

                
                Start_CANape(dest_folder)
                time.sleep(2)

                flag_start = 0
                while(flag_start == 0):
                    
                    syncFileRead = open(pathTextFile,'r')

                    valueRead = syncFileRead.read()
                    syncFileRead.close()

                    if (valueRead == '7'):
                        time.sleep(1);
                        flag_start = 1                    
                Layout = FilePath2 + "\\" +"side.lay"
                ICC_Cancel_progress=ICC_Cancel_progress+1
                ICC_Cancel_progressbar["value"] = ICC_Cancel_progress 
####                        
####                        print Layout
####                        Instrumentation().Layouts.Open(Layout)
##                       # Instrumentation().Layouts.Save(Layout)
####                        Instrumentation().ActiveLayout.Normalize()
                #Instrumentation().Layouts.Item(Layout).Activate()  
                #Instrumentation().ActiveLayout.Maximize()
                ##Instrumentation().ActiveLayout.Normalize()                  
                ADAS_HILS_AUTOMATION (SignalData,SigInfo,sig_data_sheet,myAppl)                    # This function will execute test procedure

                ICC_Cancel_progress=ICC_Cancel_progress+1
                ICC_Cancel_progressbar["value"] = ICC_Cancel_progress        
                time.sleep(2)
                syncFileWrite = open(pathTextFile,'w')
                sync_num = 9
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)
                syncFileWrite.close()
                time.sleep(2)
                


                flag = 0
                while(flag == 0):
                    syncFileRead = open(pathTextFile,'r')
                    valueRead = syncFileRead.read()
                    syncFileRead.close()
                    if (valueRead == '10'):
                       time.sleep(1);
                       flag = 1

                 
                Screenshot(screenshot_folder,Meter_Navi_Layout_Name,Diag_Layout_Name,Sub_TestCaseName_tree)
                ICC_Cancel_progress=ICC_Cancel_progress+1
                ICC_Cancel_progressbar["value"] = ICC_Cancel_progress    

                syncFileWrite = open(pathTextFile,'w')
                sync_num = 11
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)
                syncFileWrite.close()


                

                flag = 0
                while(flag == 0):
##                        print "while loop"
                    syncFileRead = open(pathTextFile,'r')

                    valueRead = syncFileRead.read()

                    syncFileRead.close()

                    if (valueRead == '8'):
                       time.sleep(3);
                       flag = 1
##                    Stop_CANape(dest_folder)
#_______________Remove unwanted files________________________________________
                print "dest_folder",dest_folder
                dest_files = os.listdir(dest_folder)
                for file in dest_files:
                    if file[-4:] == ".exe" or file[-4:] == ".ctf" or \
                       file[-4:] == ".cns" or file[-4:] == ".scr" or file[-2:] == ".c" or \
                       file == "Defect_Description_1.txt"  or \
                       file == "Sync.txt" :
                        os.remove(dest_folder + "\\" + file)
                try:
                    shutil.rmtree(dest_folder + "\\Test_Result_Automation_mcr")
                except:
                    pass
                try:
                    shutil.rmtree(dest_folder + "\\Screenshot_call_mcr")
                except:
                    pass


                interfacebook = xlrd.open_workbook(interface_sheet_path,formatting_info=True)
                interfacesheetName = interfacebook.sheet_by_name("VBA_Script_Run")
                if interfacesheetName.cell(7,1).value == 0:                                                         # if Provided judgement type not found in Result report log it to the HILS_Testing_Log file
                    print "Provided judgement type not found in Result report"
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('_________________Provided judgement type not found in Result report_________________')



#____________________________ write OK CA in gui_____________________________________
                book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                JudgementsheetName = book1.sheet_by_index(sheetNumber)
                if JudgementsheetName.cell(6,1).value!=1:
                    Result_Strng = 'CA'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Result of executed test case is  %s',Result_Strng)
                else:
                    Result_Strng = 'OK'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Result of executed test case is  %s',Result_Strng)

                ICC_Cancel_progress=ICC_Cancel_progress+1
                ICC_Cancel_progressbar["value"] = ICC_Cancel_progress                    
         
                Reset_functionality(TestCase_Start_Row,TestCase_End_Row,Primary_Test_sheet,ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row,Procedure_Test_sheet,myAppl)                    

#___________________________clear data in py-mscript_________________________________________                 
                book3 = copy(book1)
                JudgementsheetName = book3.get_sheet(sheetNumber)
                JudgementsheetName.write(0, 1, '')
                JudgementsheetName.write(1, 1, '')
                JudgementsheetName.write(2, 1, '')
                JudgementsheetName.write(3, 1, '')
                JudgementsheetName.write(4, 1, '')
                JudgementsheetName.write(5, 1, '')
                JudgementsheetName.write(6, 1, '')
                JudgementsheetName.write(7, 1, '')
                book3.save(Book_Master_Judgement_sheet)
##                    book3.Close()
                
                ManeuverTesting_result_entry.delete(0, END)
                ManeuverTesting_result_entry.insert(0,Result_Strng)
                #ICC_Cancel_Tree.item(curItem, text = ApplicationName + '_TEST_' + Sub_TestCaseName_tree, values = Result_Strng )
                ICC_Cancel_Tree.item(curItem, text = ApplicationName + '_TEST_' + Sub_TestCaseName_tree,values = (Enabled_Testcase_name[count],Result_Strng) )
                #ICC_Cancel_Tree.item(curItem, text = ApplicationName + '_TEST_' + Sub_TestCaseName_tree,text = Enabled_Testcase_name[count], values = Enabled_Testcase_name[count],values = Result_Strng )
                count = count + 1
                time.sleep(1)
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('Test case execution Completed %s', TestCaseName_tree)
                
                
                time.sleep(1)
           
                #ManeuverTesting_overall_progressbar["value"] = uid

            
                
                                
                
                myAppl.Variable(Power_Supply_path).Write(129)                    
                

                

                SignalData = []
                SigInfo =[]                         #clearing lists
                SigNames = []


            book2 = copy(book1)
            Interface_sheet = book2.get_sheet(sheetNumber)
            Interface_sheet.write(30, 0, '')
            Interface_sheet.write(31, 0, '')

            book2.save(Book_Master_Judgement_sheet)

            time.sleep(1)
            
 
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.info('##############  HILS Testing execution Completed  ##############')



##                    print "exception occured"
        
            ICC_Cancel_overall_progressbar["value"] = uid    
        except Exception, e:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')

                logging.exception('Test case execution stopped abrubtly')

                
    def Missing_Input(Missing_Input_Details):
        
        if Missing_Input_Details == "":
            pass
        else:
            #tkMessageBox.showwarning("Info",Missing_Input_Details)                                      # show all messages of user mistakes and exit code
            All_Process_TM = os.popen("tasklist").read()
            while "EXCEL.EXE" in All_Process_TM:
                All_Process_TM = os.popen("tasklist").read()
                os.system("taskkill /f /im EXCEL.EXE") 
            cmd = 'WMIC PROCESS get Caption,Commandline,Processid'
            proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)    
            for task in proc.stdout:        
                if "Pythonwin.exe" in task:
                    os.system('TASKKILL /F /IM PythonWin.exe')




    def Failsafe_Testing_ADAS_25():
       
        
      
        try:           
            Master_Result_Report_Path=Master_Result_Report             
            Master_CANape_Failsafe_Path=Master_CANape_ADAS25_Failsafe_Path
            
            global Primary_Test_sheet,sig_path,sig_value,sig_name,sig_delay,sig_reset,x,SignalData,SigInfo,SigNames
            global ECU_Array_list_JT2,ECU_Array_list_CHK,ECU_Array_list_MSG
            global Active_FAilsafe_Result_Folder

            global JT2_counter,JT2_counter_list
            global MSG_counter,MSG_counter_list
            global CHK_counter,CHK__counter_list
            global Application_start_Row,Application_end_Row,sig_data_sheet

            #global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DTC_string_path_temp,Read_vehicle_speed_path
            global Missing_Input_Details
            JT2_counter=0
            MSG_counter=0
            CHK_counter=0
            JT2_counter_list=[]
            MSG_counter_list=[]
            CHK_counter_list=[]
            CANID_list=[]
            COUNT_YES=[]
            COUNT_Category_value=[]
            Category_Array_list=[]
            Failsafe_CANID_list=[]

            end_test_str = 'end_set_test_case'
            start_test_str = 'start_set_test_case'
            Application_start_string="Meter_Navi_Enabled_Applications"
            Application_end_string="Meter_Navi_Enabled_Applications_End"
            Meter_Navi_Layout_str = "Meter_Navi_Layout_Name"
            Diag_layout_str="Diag_Layout_Name"
            CAN_Test_02_str="CAN_Test_02"
            CAN_Test_03_str="CAN_Test_03"
            cell_value_counter = 29

            global SigInfo_temp,SignalData1_temp,SignalData,SigNames,SigInfo,sig_info
            exec_var_dep=[]
            exec_delay=[]
            execute_Failsafe_cont=[]
            exec_push_var_dep=[]
            exec_wait_var_dep=[]
            execute_start_end_str=[]
            execute_cont_str=[]
            
            
            SigInfo=[]
            SigInfo_temp=[]
            SigNames=[]
            SignalData=[]
            SignalData1_temp=[]
            sig_info={}
            Failsafe_Application_sheet=''
            exec_var_dep = ['exec_var_dep']
            execute_Failsafe_cont=['exec_failsafe_cont']
            exec_delay = ['exec_delay']
            exec_push_var_dep = ['exec_push_var_dep']
            exec_wait_var_dep = ['exec_wait_var_dep']
            execute_start_end_str = ['exec_start_end']
            execute_cont_str = ['exec_cont']        
            Failsafe_Result_Folder=[]

            Failsafe_CANID_row=[]     #List used to store row number of "Y" in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET).       Used to insert result later in code
            Failsafe_CANID_column=[]  #List used to store column number of "Y" in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET).    Used to insert result later in code
            
            Meter_Navi_Layout_Name=''
            Diag_layout_Name=''
            

            global Power_Supply_path,myAppl

            try:
                Folders = os.listdir(Test_Sheet_Path)
                print Folders
                #if Failsafe_Enabled == 1 :
               #     Master_TestPattern_name = 'Master_TestPattern_FLS.xls'
                
                Master_TestPattern_name = 'Master_TestPattern_FLS.xls'
                print "Master_TestPattern_name",Master_TestPattern_name,VehicleName
                if VehicleName in Folders:
                    logging.info('Info :: Vehicle folder found ')
                    Folders = os.listdir(Test_Sheet_Path+'\\'+VehicleName)
                    print Folders
                    if Master_TestPattern_name in Folders:
                        logging.info('Info :: Test pattern sheet found ')
                        Book_Master_TP = Org_Path + "03_Master_Test_Sheet" + "\\" + VehicleName+"\\" + Master_TestPattern_name
                        book_TP = xlrd.open_workbook(Book_Master_TP,formatting_info=True)													# Open Judgement sheet
                        sheetNames_TP = book_TP.sheet_names()
                        if sig_data_sheet_str in sheetNames_TP:
                            print "Signal Data sheet exist"
                            Signal_Data_sheet = book_TP.sheet_by_name(sig_data_sheet_str)
                            End_row = Signal_Data_sheet.nrows
                            print "End_row",End_row
                            Power_Supply_path_Found = 0
                            CAR_SLCT_NO_path_Found = 0
                            DIAG_CMD_NO_path_Found = 0
                            DIAG_CMD_NO_3_path_Found = 0
                            DTC_string_path_temp_Found = 0
                            DTC_String_path_Found = 0
                            DTC_String_1_path_Found = 0
                            Read_vehicle_speed_path_Found = 0
                            for k in range(0, End_row):
                                print Signal_Data_sheet.cell(k,0).value
                                if Signal_Data_sheet.cell(k,0).value == "Power_Supply":
                                    Power_Supply_path_Found = 1
                                    Power_Supply_path = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "CAR_SLCT_NO":
                                    CAR_SLCT_NO_path_Found = 1
                                    CAR_SLCT_NO_path = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DIAG_CMD_NO_2":
                                    DIAG_CMD_NO_path_Found = 1
                                    DIAG_CMD_NO_path = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DIAG_CMD_NO_3":
                                    DIAG_CMD_NO_3_path_Found = 1
                                    DIAG_CMD_NO_path_3 = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DTC_String":
                                    DTC_String_path_Found = 1
                                    DTC_string_temp = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DTC_String_1":
                                    DTC_String_1_path_Found = 1
                                    DTC_string_1_temp = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DTC_string_ADAS2":
                                    DTC_string_path_temp_Found = 1
                                    DTC_string_path_temp = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "Read_vehicle_speed":
                                    Read_vehicle_speed_path_Found = 1
                                    Read_vehicle_speed_path = Signal_Data_sheet.cell(k, 1).value
                                else:
                                    continue
                                
                            if Power_Supply_path_Found == 1:
                                logging.info('Info :: Power_Supply_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: Power_Supply_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. Power_Supply_path not found in Test pattern sheet\n'
                            if CAR_SLCT_NO_path_Found == 1:
                                logging.info('Info :: CAR_SLCT_NO_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: CAR_SLCT_NO_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. CAR_SLCT_NO_path not found in Test pattern sheet\n'
                                
                            if DIAG_CMD_NO_path_Found == 1:
                                logging.info('Info :: DIAG_CMD_NO_path_2 found in Test pattern sheet')
                            else:
                                logging.info('Info :: DIAG_CMD_NO_path_2 not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. DIAG_CMD_NO_path not found in Test pattern sheet\n'
                            if DIAG_CMD_NO_3_path_Found == 1:
                                logging.info('Info :: DIAG_CMD_NO_3_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: DIAG_CMD_NO_3_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details +=str(error_Count) + '. DIAG_CMD_NO_3_path not found in Test pattern sheet\n'

                                
##                            if DTC_String_path_Found == 1:
##                                logging.info('Info :: DTC_String_path found in Test pattern sheet')
##                            else:
##                                logging.info('Info :: DTC_String_path not found in Test pattern sheet')
##                                error_Count = error_Count + 1
##                                Missing_Input_Details += str(error_Count) + '. DTC_String_path not found in Test pattern sheet\n'
##                                
##                                
##                            if DTC_String_1_path_Found == 1:
##                                logging.info('Info :: DTC_String_1_path found in Test pattern sheet')
##                            else:
##                                logging.info('Info :: DTC_String_1_path not found in Test pattern sheet')
##                                error_Count = error_Count + 1
##                                Missing_Input_Details += str(error_Count) + '. DTC_String_1_path not found in Test pattern sheet\n'
                                
                            if DTC_string_path_temp_Found == 1:
                                logging.info('Info :: DTC_string_path_temp found in Test pattern sheet')
                            else:
                                logging.info('Info :: DTC_string_path_temp not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. DTC_string_path_temp not found in Test pattern sheet\n'
                            if Read_vehicle_speed_path_Found == 1:
                                logging.info('Info :: Read_vehicle_speed_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: Read_vehicle_speed_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. Read_vehicle_speed_path not found in Test pattern sheet\n'
                        else:
                            logging.info('Info :: Signal Data sheet not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. Signal Data sheet not found in Test pattern sheet\n'
                    else:
                        print "Test pattern sheet not found"
                        logging.info('Info :: Test pattern sheet not found ')
                        error_Count = error_Count + 1
                        Missing_Input_Details += str(error_Count) + '. Test pattern sheet not found\n'
                else:                                                                                                               # if vehicle or keywords not found it will display message
                    print "Vehicle not found master test sheet folder"
                    logging.info('Info :: Vehicle not found master test sheet folder ')
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + '. Vehicle folder not found in master test sheet folder\n'
                    #tkMessageBox.Showinfo("Info","Vehicle not found in Master Test Sheet")
            except:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='folder not present')                        
                pass




            Test_Sheet = Book_Master_TP
            Book_Master_Judgement_sheet = Org_Path + "06_Master_Judgement_Sheet" + "\\" + "Master_Judgement_Sheet_FLS.xls"

            try:
                book = xlrd.open_workbook(Test_Sheet)                                                                       # Opening the test case sheet
                sig_data_sheet = book.sheet_by_name(sig_data_sheet_str)
            except Exception, e:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
                logging.exception('Master Test Pattern not found')            

            Book_Master_Test_sheet = xlrd.open_workbook(Test_Sheet)
            Primary_Test_sheet = Book_Master_Test_sheet.sheet_by_name("Failsafe_Primary")
            Master_Execution_sheet_End_row = Primary_Test_sheet.nrows
            CANID_Test_sheet = Book_Master_Test_sheet.sheet_by_name("CANID_LIST")

            logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                            format='%(asctime)s - %(levelname)s - %(message)s')             
            logging.info('Failsafe ADAS_2.5 Excution Started')
    #************ Screenshot function**************#
    # Brings forward Control Desk application. Takes screenshot of Meter Navi.layout and Diag layout.#			

            def Screenshot(screenshot_path,Meter_Navi_Layout_Name,Diag_layout_Name,Sub_TestCaseName_tree):

                time.sleep(1)
                try:      
                    wildcard = ".*ControlDesk Developer Version*"
                    cW = cWindow()
                    handle_manager=cW.find_window_wildcard(wildcard)
                    cW.Maximize()
                    cW.BringToTop()
                    cW.SetAsForegroundWindow()
                    bbox = win32gui.GetWindowRect(handle_manager)

                except:
                    f = open("log.txt", "w")
                    f.write(traceback.format_exc())
                    print traceback.format_exc()

                time.sleep(1)                
                try:
                    Layout = FilePath2 + "\\" +Meter_Navi_Layout_Name
                    Instrumentation().Layouts.Item(Layout).Activate()  
                    Instrumentation().ActiveLayout.Maximize()
                    snapshot=ImageGrab.grab(bbox)            
                    snapshot.save(screenshot_path+"\\"+Sub_TestCaseName_tree+"_Meter_navi.jpg")
                    time.sleep(2)
                    Layout = FilePath2 + "\\" + Diag_Layout_Name
                    Instrumentation().Layouts.Item(Layout).Activate()  
                    Instrumentation().ActiveLayout.Maximize()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(2)
                    time.sleep(1.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)
                    snapshot=ImageGrab.grab(bbox)             	
                    snapshot.save(screenshot_path+"\\"+Sub_TestCaseName_tree+"_DTC.jpg")
                    time.sleep(2)

                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('Meter_Navi_Layout_Name,Diag_Layout_Name or DIAG_CMD_NO_path not found')




                try:      
                    wildcard = ".*CANape"
                    cW = cWindow()
                    cW.find_window_wildcard(wildcard)
                    cW.Maximize()
                    cW.BringToTop()
                    cW.SetAsForegroundWindow()
                    time.sleep(1)

                except:
                    f = open("log.txt", "w")
                    f.write(traceback.format_exc())
                   ## print traceback.format_exc()

    #************ End of Screenshot function**************#

    #************ CANID_Extraction_from_CANIDList function**************#
    # Extracts CANID data from CANID_LIST sheet in Master Test Pattern. Adds to Failsafe Tree under CAN_TEST_02 and CAN_TEST_03.#
            def CANID_Extraction_from_CANIDList(CANID_Test_sheet,Category_value,TestcaseID,Var_str):
                global JT2_counter
                global MSG_counter
                global CHK_counter
                flag1=0
                flag2=0
                flag3=0
                
                CANID_List_End_Row_str="END_CANID_List"
                CAN_Test_02_str="CAN_Test_02"
                CAN_Test_03_str="CAN_Test_03"
                
                ECU_Sensor_Column= 0
                CAN_ID_Cloumn= 1
                JT2_Column= 2
                Checksum_Coulmn= 3
                Message_Counter_Column=4
                #Enable_Disable_Column=5


                CANID_Test_sheet_End_row = CANID_Test_sheet.nrows
                CANID_Test_sheet_End_col = CANID_Test_sheet.ncols
                for k in range(0, CANID_Test_sheet_End_row):
                    if CANID_Test_sheet.cell(k,0).value== CANID_List_End_Row_str :
                        CANID_List_End_Row= k
                        break

                for k in range(0, CANID_Test_sheet_End_col):
                    if CANID_Test_sheet.cell(0,k).value== Var_str :
                        Enable_Disable_Column= k
                        break

                print "Enable_Disable_Column",Enable_Disable_Column,Var_str                    
               
                for i in range (1, CANID_List_End_Row):      
                    

                    if TestcaseID == CAN_Test_02_str:
                       
                        if (CANID_Test_sheet.cell(i,Enable_Disable_Column).value=='Y'):
                            CAN_ID=CANID_Test_sheet.cell(i,CAN_ID_Cloumn).value
                            ECU_name=CANID_Test_sheet.cell(i,ECU_Sensor_Column).value 
                            CANID_list.append(CAN_ID)

                            if (CANID_Test_sheet.cell(i,JT2_Column).value=='Y'):
                                if(flag1==0):
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["JT2"]=OrderedDict()
                                    flag1=1
                                    
                                                       
                                if ECU_name in ECU_Array_list_JT2:
                               
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["JT2"][ECU_name][CAN_ID]=OrderedDict()
                                  
                                    
                                else :
                               
                                    ECU_Array_list_JT2.append(ECU_name)
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["JT2"][ECU_name]=OrderedDict()
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["JT2"][ECU_name][CAN_ID]=OrderedDict()
                                  

                                JT2_counter = JT2_counter + 1
                             
                        
                                                 
                               
                    if TestcaseID == CAN_Test_03_str:
                        
                        if (CANID_Test_sheet.cell(i,Enable_Disable_Column).value=='Y'):
                            if flag2==0:
                                Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Message Counter"]=OrderedDict()
                                flag2=1
                                
                       
                            ECU_name=CANID_Test_sheet.cell(i,ECU_Sensor_Column).value
                            CAN_ID=CANID_Test_sheet.cell(i,CAN_ID_Cloumn).value
                            if (CANID_Test_sheet.cell(i,Message_Counter_Column).value=='Y'):
                                
                                if ECU_name in ECU_Array_list_MSG:
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Message Counter"][ECU_name][CAN_ID]=OrderedDict()
                                               
                                else :
                                    ECU_Array_list_MSG.append(ECU_name)
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Message Counter"][ECU_name]=OrderedDict()
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Message Counter"][ECU_name][CAN_ID]=OrderedDict()

                                MSG_counter = MSG_counter + 1

                            if (CANID_Test_sheet.cell(i,Checksum_Coulmn).value=='Y'):
                                if flag3==0:
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Checksum"]=OrderedDict()

                                    flag3=1                                
                                
                                if ECU_name in ECU_Array_list_CHK:
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Checksum"][ECU_name][CAN_ID]=OrderedDict()
                                               
                                else :
                                    ECU_Array_list_CHK.append(ECU_name)
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Checksum"][ECU_name]=OrderedDict()
                                    Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]["Checksum"][ECU_name][CAN_ID]=OrderedDict()
                                
                                CHK_counter = CHK_counter + 1






                return Failsafe_Dict   

    #************ End of CANID_Extraction_from_CANIDList function**************#				
                
    #************ Make_SignalData_Array function**************#
    # Makes SignalData Array to be given to ADAS_HILS_AUTOMATION. This contains the extracted procedure for a particular test case #

            def Make_SignalData_Array(Procedure_Test_sheet,x):
                global SignalData1_temp,SigInfo_temp,sig_info
                SigNames=[]
                sig_name=None
                sig_delay=None
                sig_path=None
                sig_value=None
                sig_reset=None
                
                sig_name = Procedure_Test_sheet.cell(x,0).value
                sig_delay = Procedure_Test_sheet.cell(x,2).value   
                sig_path = Procedure_Test_sheet.cell(x, 6).value
                sig_value = Procedure_Test_sheet.cell(x,7).value
                sig_reset = Procedure_Test_sheet.cell(x,8).value
                SigNames.append(sig_name)    
                sig_data = [sig_path,sig_value,sig_reset]                  
                sig_info[sig_name] = sig_data
                SigInfo_temp.append(sig_info[sig_name])
                SignalData1_temp =data_function(Procedure_Test_sheet,sig_path,sig_value,sig_name,sig_delay,sig_reset,x)
                return SigInfo_temp,SignalData1_temp

    #************ End of Make_SignalData_Array function**************#	

    #************ Execute_TestCase function**************#
    # Calls CANape amd ADASS_HILS_Automation. #
            def Execute_TestCase(dest_Failsafe_Result_Folder,SignalData1,SigInfo,sig_data_sheet,Screenshot_Name,curItem_failsafe,CANID,Meter_Navi_Layout_Name,Diag_layout_Name,Sub_TestCaseName_tree):
                
                pathTextFile = dest_Failsafe_Result_Folder
                pathTextFile = pathTextFile + "\\" + "Sync.txt"                

                screenshot_Failsafe_Result_Folder = dest_Failsafe_Result_Folder + "\\"+"00_Screenshot"
                os.mkdir(screenshot_Failsafe_Result_Folder)                

                write_ss_info = open(dest_Failsafe_Result_Folder + "\\Screen_Shot.txt","w")
                screenshot_path=dest_Failsafe_Result_Folder + "\\" + "00_Screenshot"
                write_ss_info.write(dest_Failsafe_Result_Folder + "\\" + "00_Screenshot\n")
                write_ss_info.write(Screenshot_Name)           
                write_ss_info.close()
                #myAppl.Variable(Power_Supply_path).Write(1)


                time.sleep(2)
                myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)
                time.sleep(1)


                myAppl.Variable("simState").Write(0)                                                                    # 'Reset' Simstate
                time.sleep(.5)
                myAppl.Variable("simState").Write(2)
                
                myAppl.Variable(Power_Supply_path).Write(129)
                time.sleep(1)
                myAppl.Variable(Power_Supply_path).Write(1)                                                                
                time.sleep(6)
                
                myAppl.Variable(DIAG_CMD_NO_path).Write(3)                                                          #Clear DTC
                time.sleep(.5)                  
                myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                time.sleep(.5)
         
                flag_speed = 0
                count = 0
                while(flag_speed == 0):                                                                             #check for vehicle speed and DTC 
                    print "DIAG_CMD_NO_path",DIAG_CMD_NO_path
                    read_velocity_1 = myAppl.Variable(Read_vehicle_speed_path).Read()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(2)                                                          #Read DTC
                    time.sleep(0.5)                 
                    Actual_DTC_set = myAppl.Variable(DTC_string_path_temp).Read()
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(.5)
                    print "Actual_DTC_set",int(float(Actual_DTC_set))
                    time.sleep(7)
                    read_velocity_2 = myAppl.Variable(Read_vehicle_speed_path).Read()
                    print "vehicle speed read",int(float(read_velocity_2)),int(float(read_velocity_1))
                    if (int(float(read_velocity_2)) - int(float(read_velocity_1)) > 0) :							#and (int(float(Actual_DTC_set)) == 0) :                    
                        flag_speed = 1
                    else:
                        myAppl.Variable(Power_Supply_path).Write(129)
                        time.sleep(1)
                        myAppl.Variable(Power_Supply_path).Write(1)                                                                 # 'Set' Simstate
                        time.sleep(6)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(3)
                        time.sleep(.5)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                        time.sleep(.5)                        
                        myAppl.Variable(DIAG_CMD_NO_path).Write(2)
                        time.sleep(0.5)
                        myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                        time.sleep(0.5)

                    if count > 5:    
                        flag_speed = 1
                    count = count + 1 

#               if count > 5:
#                    Controldesk_Load_Reload(Write_Var)    


                Start_CANape(dest_Failsafe_Result_Folder)
                time.sleep(2)
                

                flag_start = 0
                while(flag_start == 0):
                  
                    syncFileRead = open(pathTextFile,'r')

                    valueRead = syncFileRead.read()
                    syncFileRead.close()

                    if (valueRead == '7'):
                        time.sleep(1);
                        flag_start = 1                    


                ADAS_HILS_AUTOMATION (SignalData1,SigInfo,sig_data_sheet,myAppl)
           
                time.sleep(2)
                syncFileWrite = open(pathTextFile,'w')
                sync_num = 9
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)
                syncFileWrite.close()
                time.sleep(2)
                


                flag = 0
                while(flag == 0):
                    syncFileRead = open(pathTextFile,'r')
                    valueRead = syncFileRead.read()
                    syncFileRead.close()
                    if (valueRead == '10'):
                       time.sleep(1);
                       flag = 1

                Screenshot(screenshot_path,Meter_Navi_Layout_Name,Diag_Layout_Name,Sub_TestCaseName_tree)


                syncFileWrite = open(pathTextFile,'w')
                sync_num = 11
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)
                syncFileWrite.close()
                

                flag = 0
                while(flag == 0):

                    syncFileRead = open(pathTextFile,'r')

                    valueRead = syncFileRead.read()

                    syncFileRead.close()

                    if (valueRead == '8'):
                       time.sleep(1);
                       flag = 1



                dest_files = os.listdir(dest_Failsafe_Result_Folder)
                for file in dest_files:
                    #if file[-4:] == ".exe" or file[-4:] == ".ctf" or \
                    if    file[-4:] == ".cns" or file[-4:] == ".scr" or file[-2:] == ".c" or \
                       file == "Defect_Description11.txt"  or \
                       file == "Sync.txt" :
                        os.remove(dest_Failsafe_Result_Folder + "\\" + file)
                try:
                    shutil.rmtree(dest_Failsafe_Result_Folder + "\\Test_Result_Automation_mcr")
                except:
                    pass
                try:
                    shutil.rmtree(dest_Failsafe_Result_Folder + "\\Screenshot_call_mcr")
                except:
                    pass


                sheetName ="Py-MScript"                        
                book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                JudgementsheetName = book1.sheet_by_name(sheetName)
                if JudgementsheetName.cell(6,1).value!=1:
                    Result_Strng = 'CA'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Result of executed test case is  %s',Result_Strng)
                else:
                    Result_Strng = 'OK'
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Result of executed test case is  %s',Result_Strng)


                print "interface_sheet_path",interface_sheet_path
                interfacebook = xlrd.open_workbook(interface_sheet_path,formatting_info=True)
                interfacesheetName = interfacebook.sheet_by_name("VBA_Script_Run")
                print "interfacesheetName.cell(7,1).value",interfacesheetName.cell(7,1).value,interfacesheetName.cell(4,1).value
                if interfacesheetName.cell(7,1).value == 0:                                                         # if Provided judgement type not found in Result report log it to the HILS_Testing_Log file
                    print "Provided judgement type not found in Result report"
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('_________________Provided judgement type not found in Result report_________________')
                    

                FAILSAFE_result_entry["state"] = NORMAL
                FAILSAFE_result_entry.delete(0, END)
                FAILSAFE_result_entry.insert(0,Result_Strng)
                FAILSAFE_result_entry["state"] = DISABLED
                time.sleep(3)
                FAILSAFE_result_entry["state"] = NORMAL
                FAILSAFE_result_entry.delete(0, END)
                FAILSAFE_result_entry["state"] = DISABLED
                Failsafe_Tree.item(curItem_failsafe, text = CANID, values =Result_Strng )
                
     


                
                SignalData = []
                SigInfo =[]
                SigNames = []
                SigInfo_temp=[]
                SignalData1_temp=[]

    #************ End of Execute_TestCase function**************#	

    #************ Execute_TestCase function**************#
    # Reset Signals and clear DTC. #

            def Reset_functionality(TestCase_Start_Row,TestCase_End_Row,Primary_Test_sheet,ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row,Failsafe_Application_sheet,myAppl):
                try:
                    for m in range(ApplicationSheet_TestCase_Start_Row+1,ApplicationSheet_TestCase_End_Row ):                
                        if Failsafe_Application_sheet.cell(m,6).value!='' :
                           
                            set_sig_reset_value = None
                            set_sig_path = None
                            set_sig_default_value = None
                            set_sig_Appl = None
                            set_sig_reset_value = int(Failsafe_Application_sheet.cell(m,8).value)                    
                            set_sig_path = Failsafe_Application_sheet.cell(m,6).value
                            set_sig_default_value = Failsafe_Application_sheet.cell(m,7).value                   
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(0.5)
                            

                    for m in range(TestCase_Start_Row + 1,TestCase_End_Row):
                       
                        if Primary_Test_sheet.cell(m,6).value != '' :
                           
                            set_sig_reset_value = None
                            set_sig_path = None
                            set_sig_default_value = None
                            set_sig_Appl = None
                            set_sig_reset_value = int(Primary_Test_sheet.cell(m,8).value)
                            set_sig_path = Primary_Test_sheet.cell(m,6).value
                            set_sig_default_value = Primary_Test_sheet.cell(m,7).value
                            myAppl.Variable(set_sig_path).Write(0)
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)                            
                            time.sleep(0.5)
                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('signal Path , reset value not found in sheet')
             
                try:   

                    myAppl.Variable(Power_Supply_path).Write(129)
                    time.sleep(2)
                    myAppl.Variable(Power_Supply_path).Write(1)
                    time.sleep(2)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(3)
                    time.sleep(1.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)
                    myAppl.Variable(Power_Supply_path).Write(129)
                    time.sleep(1)
                    myAppl.Variable(Power_Supply_path).Write(1)
                    time.sleep(1)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(2)
                    time.sleep(0.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(1)

                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                        level=logging.INFO,
                                        format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('power supply & Diag Paths are wrong....... Check signal paths in signal data sheet')
             
    #************ End of Reset_functionality function**************#


            def CAN_Message_Replacement(CAN_Value_1,CAN_Value_2):
                book1 = xlrd.open_workbook(str(Book_Master_Judgement_sheet),formatting_info=True)
                sheetNames = book1.sheet_names()
                sheetNumber = 0
                for i in sheetNames:
                    if "Py-MScript" in i:
                        break
                    sheetNumber = sheetNumber + 1
                        
                book2 = copy(book1)
                Interface_sheet = book2.get_sheet(sheetNumber)
                Interface_sheet.write(25, 2, CAN_Value_1)
                Interface_sheet.write(25, 3, CAN_Value_2)
           
                book2.save(Book_Master_Judgement_sheet)
                


    #************ Copy_TestCase_Data_to_PyMScript function**************#
    # Copy information like Vehicle details, Test Case details to Py-MScript sheet in Master Judgement sheet. #
                
            def Copy_TestCase_Data_to_PyMScript(Write_Var,ApplicationName_tree,Sub_TestCaseName_tree,ECU,CANID,dest_Failsafe_Result_Folder):
                

               
                book1 = xlrd.open_workbook(str(Book_Master_Judgement_sheet),formatting_info=True)
                sheetNames = book1.sheet_names()
                sheetNumber = 0
                for i in sheetNames:
                    if "Py-MScript" in i:
                        break
                    sheetNumber = sheetNumber + 1
                        
                book2 = copy(book1)
                Interface_sheet = book2.get_sheet(sheetNumber)
                Interface_sheet.write(0, 1, VehicleName)
           
                Interface_sheet.write(1, 1, RegionName)
                
                Interface_sheet.write(2, 1, Write_Var)
              
                Interface_sheet.write(3, 1, ApplicationName_tree)
           
                if ApplicationName_tree=="CAN":
                    Interface_sheet.write(4, 1, ECU+"_"+Sub_TestCaseName_tree)
                if ApplicationName_tree!="CAN":
                    Interface_sheet.write(4, 1,Sub_TestCaseName_tree)
                Interface_sheet.write(22, 1, 1)
                Interface_sheet.write(23, 1, Sub_TestCaseName_tree)
                Interface_sheet.write(24, 1, ECU)
                Interface_sheet.write(25, 1, CANID)           
                Interface_sheet.write(8, 1, dest_Failsafe_Result_Folder)
                book2.save(Book_Master_Judgement_sheet)
                print "part1"
                xlapp1 = win32com.client.Dispatch("Excel.Application")   #To open Excel
                if os.path.exists(str(interface_sheet_path)):
                 
                    xlapp1.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    print "part2"
                    if ApplicationName_tree=="CAN":
                        xlapp1.Application.Run("Interface_VBA.xls!module9.replace_data")
                    else:
                        xlapp1.Application.Run("Interface_VBA.xls!module13.replace_data_WithoutCAN")
            
                   # xlapp1.Workbooks.Close()
                    print "part3"
                    xlapp1.Application.Run("Interface_VBA.xls!module11.Make_Canape_failsafe")
                  

                    xlapp1.Workbooks.Close()                
                
            def data_function(Primary_Test_sheet,sig_path,sig_value,sig_name,sig_delay,sig_reset,x):
                signal_data = []
                save_var = None
                if str(Primary_Test_sheet.cell(x,1).value) in \
                   execute_start_end_str:                                                                       # If string in cloumn 2 is 'exec_start_end' then call 'get_data_execute_start_end_str' to collect test case data 
                    signal_data = get_data_execute_start_end_str(sig_name,
                                                                 sig_path,
                                                                 Primary_Test_sheet,
                                                                 sig_delay,x,
                                                                 sig_data_sheet)
                    save_var = 1
                elif str(Primary_Test_sheet.cell(x,1).value) in execute_cont_str:                                # If string in cloumn 2 is 'exec_cont' then call 'get_data_execute_conti' to collect test case data
                                  
                    signal_data = get_data_execute_cont(sig_name,sig_path,
                                                        Primary_Test_sheet,
                                                        sig_delay,x,
                                                        sig_data_sheet)
                    save_var = 1


                elif str(Primary_Test_sheet.cell(x,1).value) in execute_Failsafe_cont:                                # If string in cloumn 2 is 'exec_cont' then call 'get_data_execute_conti' to collect test case data
                    

                    Failsafe_Application_sheet_row=x

                  
                    signal_data = get_data_execute_Failsafe_cont(sig_name,sig_path,
                                                        Failsafe_Application_sheet,
                                                        sig_delay,Failsafe_Application_sheet_row,
                                                        sig_data_sheet)
                    save_var = 1

                
                   
                elif str(Primary_Test_sheet.cell(x,1).value) in exec_var_dep:                                    # If string in cloumn 2 is 'exec_var_dep' then call 'get_data_execute_var_dep' to collect test case data   
                    signal_data = get_data_execute_var_dep(sig_name,sig_path,
                                                           Primary_Test_sheet,
                                                           sig_delay,x,
                                                           sig_data_sheet)
                    save_var = 1
                    
                elif str(Primary_Test_sheet.cell(x,1).value) in exec_delay:                                      # If string in cloumn 2 is 'exec_delay' then call 'get_execute_delay' to collect test case data  
                    signal_data = get_execute_delay(sig_name,sig_delay)
                    save_var = 1



                SignalData.append(signal_data) 
            


                if save_var == 1:                                                                               # Save_var = 1 signifies that atleast 1 signal is present in that test case 
                    save_var = 0
                    
                signal_data = []
                sig_info = {}
                return SignalData

            

            
            for k in range(0, Master_Execution_sheet_End_row):
                    if Primary_Test_sheet.cell(k,0).value== start_test_str :
                            TestCase_Start_Row= k
                            break

            
            for k in range(0, Master_Execution_sheet_End_row):
                if Primary_Test_sheet.cell(k,0).value== end_test_str:
                        TestCase_End_Row= k
                        break

                for k in range(0, Master_Execution_sheet_End_row):
                    if Primary_Test_sheet.cell(k,0).value== Meter_Navi_Layout_str :
                        Meter_Navi_Layout_Name= Primary_Test_sheet.cell(k,1).value
                        break

                for k in range(0, Master_Execution_sheet_End_row):
                    if Primary_Test_sheet.cell(k,0).value== Diag_layout_str :
                        Diag_Layout_Name= Primary_Test_sheet.cell(k,1).value
                        break
            
            Vehicle_Details = VehicleName +  '_' + RegionName
            Vehicle_Name = VehicleName + '_' + RegionName + '_' + PartNo
            Actual_PartNo=PartNo.split("_")[1]




            Failsafe_Dict = OrderedDict()                    # This makes dictionary required for Busoff Tree
            Failsafe_Dict[Vehicle_Details]= OrderedDict()
            Failsafe_Dict[Vehicle_Details][PartNo]= OrderedDict()
            Active_Test = 'Failsafe'       
            Failsafe_Dict[Vehicle_Details][PartNo][Active_Test]=OrderedDict()
            Failsafe_Result_Folder=[]
            Category_Array_list=[]
            Failsafe_CANID_list=[]
            Failsafe_CANID_row=[]     
            Failsafe_CANID_column=[]  
            CANID_list=[]

            DIMPSheet_Failsafe_CANID_List = WorkBook.sheet_by_name("Failsafe")          #This opens the SECOND Sheet of DISPATCH Sheet i.e.Sheet containing list of active Message Counter CANID
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
            logging.info('Dispatch Sheet for Failsafe ADAS_2.5 Loaded')
            DIMPSheet_Failsafe_CANID_List_Col = DIMPSheet_Failsafe_CANID_List.ncols
            DIMPSheet_Failsafe_CANID_List_Row = DIMPSheet_Failsafe_CANID_List.nrows


            Category_Dictionary_for_All_Variants={}
            Current_Variant_ECU_Array_list_JT2={}
            Current_Variant_ECU_Array_list_CHK={}
            Current_Variant_ECU_Array_list_MSG={}
            application = []
            for k in range(0,len(Variant)):

                Category_Array_list=[]
                CANID_list=[]
                ECU_Array_list_JT2=[]
                ECU_Array_list_CHK=[]
                ECU_Array_list_MSG=[]



                JT2_counter=0
                MSG_counter=0
                CHK_counter=0
                counter = 0
                data  = Variant[k].split('_')
                Variant_Number_string = data[2]
                VariantName_tree = ' Variant ' + data[2]
                print VariantName_tree
                Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree]=OrderedDict()
                Failsafe_Result_Folder.append(VehicleNameFolder[k]+"\\01_Failsafe")             
                

                



                Var_str =  'Variant_0' +  str(k +1)    #This is used to obtain Variant number in format "Variant_0X" eg. Variant_01 Variant_02 etc

                Var_Row = 0    
                Var_col = 0   
                CANID_row=0     
                CANID_col=0     


                for i in range (0,  DIMPSheet_Failsafe_CANID_List_Row):             # Loop for traversing through the EXCEL sheet
                    for j in range(0,DIMPSheet_Failsafe_CANID_List_Col):
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == Var_str:         #This finds row and column of the particular variant
                            Var_col =  j
                            Var_Row = i

                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == Application_start_string:         #This finds row and column of the Meter_Navi_Enabled_Applications
                            Application_start_col =  j
                            Application_start_Row = i

                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == Application_end_string:         #This finds row and column of the Meter_Navi_Enabled_Applications_End
                            Application_end_col =  j
                            Application_end_Row = i
                            
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == 'Test_Case':        #This finds row and column of Test Cases
                            CANID_col =  j
                            CANID_row = i
                       
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == 'Category':        #This finds row and column of Category
                            CAN_Channel_col =  j
                            CAN_Channel_row = i
                      
                        else:
                            continue
                        
                book3 = xlrd.open_workbook(str(Book_Master_Judgement_sheet),formatting_info=True)
                sheetNames = book3.sheet_names()
                sheetNumber = 0
                for i in sheetNames:
                    if "Py-MScript" in i:
                        break
                    sheetNumber = sheetNumber + 1
            #                    
                book4 = copy(book3)

                        
                Pymscript_sheet = book4.get_sheet(sheetNumber)
                for i in range (Application_start_Row,  Application_end_Row+1):             # Loop for traversing through the EXCEL sheet
                  
                
                    Appl_name=DIMPSheet_Failsafe_CANID_List.cell(i,0).value
                    Enb_Appl=DIMPSheet_Failsafe_CANID_List.cell(i,1).value
                    
            
                    Pymscript_sheet.write(cell_value_counter+1,0, Appl_name)
                    Pymscript_sheet.write(cell_value_counter+1,1, Enb_Appl)
                    cell_value_counter=cell_value_counter+1
                    book4.save(Book_Master_Judgement_sheet)
               ## book4.close(Book_Master_Judgement_sheet)  

             


                for i in range (Var_Row + 1, DIMPSheet_Failsafe_CANID_List_Row):      # ( Var_Row +2 ) contains the string "Y" or "N"

                    if (DIMPSheet_Failsafe_CANID_List.cell(i,Var_col).value=='Y'):
                     
                                   #This counts the number of 'Y' for a particular variant.
                        Category_value=DIMPSheet_Failsafe_CANID_List.cell(i,CAN_Channel_col).value
                        TestcaseID=DIMPSheet_Failsafe_CANID_List.cell(i,CANID_col).value    #This extracts the CANID form the sheet
                        Failsafe_CANID_list.append(DIMPSheet_Failsafe_CANID_List.cell(i,CANID_col).value)   #This adds CANID to a list
                        Failsafe_CANID_row.append(i)   #This stores the row number of all "Y" for future reference
                    
                  
        

                        if (Category_value in Category_Array_list) and (Category_value=='CAN'):
                           

                            Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]=OrderedDict()
                            if TestcaseID==CAN_Test_03_str:
                                
                                CANID_Extraction_from_CANIDList(CANID_Test_sheet,Category_value,TestcaseID,Var_str) 

                            
                        elif Category_value in Category_Array_list:
                            Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]=OrderedDict()
                            counter=counter+1
                            
                            
                            

                        
                        else :

                            Category_Array_list.append(Category_value)
                            Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value]=OrderedDict()
                            Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][Category_value][TestcaseID]=OrderedDict()
                            counter = counter +1
                         
                            if TestcaseID==CAN_Test_02_str:
                                CANID_Extraction_from_CANIDList(CANID_Test_sheet,Category_value,TestcaseID,Var_str)
                                counter=counter-1
                                
                               
        

      
                                                         
                        
    #failsafe_errorha
                for key,value in Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree].iteritems():
                    print "key",key
                    if not key in application:
                        application.append(key)

            try:                
                print "application",application
                error_Count = 0
                check = os.path.exists(Book_Master_Judgement_sheet)                
                if (check == True):
                    book_JS = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)													# Open Judgement sheet
                    sheetNames_JS = book_JS.sheet_names()
                    sheetNumber_JS = 0           
                    print "sheetNames_JS",sheetNames_JS
                    for j in range(0,len(application)):
                        Sheet_No = 0
                        print "application[j]",application[j]
                        for i in sheetNames_JS: 																							# Loop for finding Application sheet
                            print "i",i
                            if application[j] == i:
                                Sheet_No = 1
                                break
                            sheetNumber_JS = sheetNumber_JS + 1                                                
                        if Sheet_No == 0:
                            error_Count = error_Count + 1
                            logging.info('Info ::'+ '________________________' + str(error_Count)+ ". " + str(application[j]) + ' application sheet not found in Judgement sheet for '+VehicleName)                   
                            Missing_Input_Details += str(error_Count) + ". " + str(application[j]) + ' application sheet not found in Judgement sheet \n'              
                else:
                    error_Count = error_Count + 1
                    logging.info('Info ::'  + '________________________' + str(error_Count)+". " +' Master_Judgement_Sheet_FLS not present for '+VehicleName)
                    Missing_Input_Details += str(error_Count) + ". " + ' Master_Judgement_Sheet_FLS not present for ' +"\n"

            except:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='Master_Judgement_Sheet_FLS not present')                        
                pass

            try:
                check = os.path.exists(Test_Sheet)                
                if (check == True):
                    book_TP = xlrd.open_workbook(Test_Sheet,formatting_info=True)													# Open Judgement sheet
                    sheetNames_TP = book_TP.sheet_names()
                    sheetNumber_TP = 0
                    print "sheetNames_TP",sheetNames_TP
                    for j in range(0,len(application)):
                        Sheet_No = 0
                        print "application[j]",application[j]
                        for i in sheetNames_TP: 																							# Loop for finding Application sheet
                            print "i",i
                            if application[j] == i:
                                Sheet_No = 1
                                break
                            sheetNumber_TP = sheetNumber_TP + 1

                          
                        if Sheet_No == 0:
                            error_Count = error_Count + 1
                            logging.info('Info ::' + '________________________' + str(error_Count)+ ". " + str(application[j]) + ' application sheet not found in testpattern sheet for '+VehicleName)                   
                            Missing_Input_Details += str(error_Count) + ". " + str(application[j]) + ' application sheet not found in testpattern sheet \n'              
                else:
                    error_Count = error_Count + 1
                    logging.info('Info ::' + '________________________'+str(error_Count) +". "+ ' Master_TestPattern_FLS not present for '+VehicleName)            
                    Missing_Input_Details += str(error_Count) + ". " + ' Master_TestPattern_FLS not present for ' +"\n"
            except:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='Master_TestPattern_FLS not present')                        
                pass
                


            try:
                for j in range(0,len(application)):
                    Folders = os.listdir(Master_Result_Report)
                    if application[j] in Folders:
                        Folders = os.listdir(Master_Result_Report+'\\'+application[j])
                        if application[j] +'.xls' in Folders:                                                  #checks for master result workbook of application
                            logging.info('Info :: Result sheet found ')
                        else:
                            print "Result sheet not found"
                            error_Count = error_Count + 1
                            logging.info('Info ::' +'________________________' + str(error_Count) +". "+application[j]+'Result Report not found ')              
                            Missing_Input_Details += str(error_Count) + '. '+str(application[j])+" Master Result Report not found\n"
                    else:
                        print " Application not found master test sheet folder"
                        error_Count = error_Count + 1
                        logging.info('Info ::________________________' + str(error_Count) + '. '+ str(application[j]) +'folder not found in master result report folder ')              
                        Missing_Input_Details += str(error_Count) + '. '+str(application[j]) +" folder not found in master result report folder\n"

            except:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='folder not present')                        
                pass

            
            print "Missing_Input_Details",Missing_Input_Details
            if Missing_Input_Details == '':
                pass
            else:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='execution stop due to above details not present')           


            #Missing_Input(str(Missing_Input_Details))





            

            uid_MSG_prev=uid      
            Failsafe_Tree = construct_JSON_tree(Failsafe_Dict,frame7)
            uid_MSG_prev=0
            curItem = 0
            Var_Val=0
            FAILSAFE_vehicle_id_entry["state"] = NORMAL
            FAILSAFE_vehicle_id_entry.delete(0, END)
            FAILSAFE_vehicle_id_entry.insert(0, Vehicle_Details)
            FAILSAFE_vehicle_id_entry["state"] = DISABLED
            FAILSAFE_overall_progressbar["maximum"]=uid           
            FAILSAFE_overall_progressbar["value"] = 0       

            for j in range(0, uid):
               
                    
              
                overall_progressbar_uid=j
                FAILSAFE_overall_progressbar["value"] = overall_progressbar_uid
                Ances_array=[]
                TestCaseNameId = ''
                ApplicationNameId = ''
                TestCaseNameId = ''
                CAN_CategoryId=''
                ECUId=''
                CANIDId=''
                VariantName_tree = ''
                ApplicationName_tree = ''
                CAN_Category_tree=''
                ECU_ID_tree=''
                CANID_tree=''



                
                FAILSAFE_result_entry.delete(0, END)
                TestCaseName_tree = ''
                FAILSAFE_CAN_ID_entry.delete(0, END)
                FAILSAFE_CAT_entry.delete(0, END)
                FAILSAFE_variant_entry.delete(0, END)
                FAILSAFE_ECU_entry["state"] = NORMAL
                FAILSAFE_ECU_entry.delete(0, END)
                curItem= curItem + 1
                Failsafe_Tree.selection_set(curItem)
                Item=curItem
              
        
                book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)
                sheetNames = book1.sheet_names()
                sheetNumber = 0

                for i in range(0, 10):
                    dest_folder = folder_path                                                 
                    ParentItem=Failsafe_Tree.parent(Item)
                    Ances_array.append(ParentItem)
                    Item=Ances_array[i]
                  
                    if Ances_array[i]=='':
                        break
                    else :
                        continue
                

                   
                Heir = len(Ances_array)



                if Heir == 6:               
                    TestCaseNameId = curItem
                    ApplicationNameId = Ances_array[0]
                    VariantNameId = Ances_array[1]
                    CAN_CategoryId=''
                    ECUId=''
                    CANIDId=''
                    
                elif Heir == 5:
                    ApplicationNameId = curItem             
                    VariantNameId = Ances_array[0]
                    TestCaseNameId = ''
                    CAN_CategoryId=''
                    ECUId=''
                    CANIDId=''
                    
                elif Heir == 4:
                    VariantNameId = curItem
                    ApplicationNameId = ''
                    TestCaseNameId = ''
                    CAN_CategoryId=''
                    ECUId=''
                    CANIDId=''

                elif Heir == 7:
                
                    CAN_CategoryId = curItem
                    TestCaseNameId =  Ances_array[0]                   
                    ApplicationNameId = Ances_array[1]
                    VariantNameId = Ances_array[2]
                 
                    CANIDId=''

                elif Heir == 8:
               
                    ECUId = curItem
                    CAN_CategoryId=Ances_array[0]       
                    TestCaseNameId =  Ances_array[1]                   
                    ApplicationNameId = Ances_array[2]
                    VariantNameId = Ances_array[3]
                
                    CANIDId=''
                    
                elif Heir == 9:

                    CANIDId = curItem
                    ECUId=Ances_array[0]
                    CAN_CategoryId=Ances_array[1]       
                    TestCaseNameId =  Ances_array[2]                   
                    ApplicationNameId = Ances_array[3]
                    VariantNameId = Ances_array[4]
          

                    
                else:
                    VariantNameId = ''
                    ApplicationNameId = ''                   
                    TestCaseNameId = ''

                if VariantNameId=='':
                    FAILSAFE_variant_entry.delete(0, END)
                else:
                    VariantName_tree = Failsafe_Tree.item(VariantNameId, 'text')
                    print "VariantName_tree",VariantName_tree                        


                    data  = VariantName_tree.split(' ')
                    print "VariantName_tree 2837",VariantName_tree
                    VariantName_tree = ' Variant ' + data[2]                    
                
                    FAILSAFE_variant_entry["state"] = NORMAL
                    FAILSAFE_variant_entry.delete(0, END)
                    FAILSAFE_variant_entry.insert(0, VariantName_tree)
                    print 'present variant is ' + VariantName_tree
                    FAILSAFE_variant_entry["state"] = DISABLED
                
                if ApplicationNameId=='':
                    FAILSAFE_CAT_entry.delete(0, END)
                else:
                    ApplicationName_tree = Failsafe_Tree.item(ApplicationNameId, 'text')
                    FAILSAFE_CAT_entry["state"] = NORMAL
                    FAILSAFE_CAT_entry.delete(0, END)
                    print 'present appli is ' + ApplicationName_tree
                    FAILSAFE_CAT_entry.insert(0, ApplicationName_tree)
                    FAILSAFE_CAT_entry["state"] = DISABLED
                  
                if TestCaseNameId == '':
           
                    FAILSAFE_CAN_ID_entry.delete(0, END)
                else:
                    TestCaseName_tree = Failsafe_Tree.item(TestCaseNameId, 'text')
                    FAILSAFE_CAN_ID_entry["state"] = NORMAL
                    FAILSAFE_CAN_ID_entry.delete(0, END)
                    FAILSAFE_CAN_ID_entry.insert(0, TestCaseName_tree)
                    print 'present test case is ' + TestCaseName_tree
                    FAILSAFE_CAN_ID_entry["state"] = DISABLED

                if CAN_CategoryId == '':
           
                    pass
                else:
                    CAN_Category_tree = Failsafe_Tree.item(CAN_CategoryId, 'text')
                    print 'present CAN_Category_tree case is ' + CAN_Category_tree
                     


                if ECUId == '':
           
                    FAILSAFE_ECU_entry.delete(0, END)
                else:
                    ECU_ID_tree = Failsafe_Tree.item(ECUId, 'text')
                    FAILSAFE_ECU_entry["state"] = NORMAL
                    FAILSAFE_ECU_entry.delete(0, END)
                    FAILSAFE_ECU_entry.insert(0, ECU_ID_tree)
                    print 'present ECU_ID_tree case is ' + ECU_ID_tree
                    FAILSAFE_ECU_entry["state"] = DISABLED



                if CANIDId == '':
           
                    FAILSAFE_CAN_ID_entry.delete(0, END)
                else:
                    CANID_tree = Failsafe_Tree.item(CANIDId, 'text')
                    FAILSAFE_CAN_ID_entry["state"] = NORMAL
                    FAILSAFE_CAN_ID_entry.delete(0, END)
                    FAILSAFE_CAN_ID_entry.insert(0, CANID_tree)
                    print 'present CANID_tree case is ' + CANID_tree
                    FAILSAFE_CAN_ID_entry["state"] = DISABLED
               

                    
                    

                if TestCaseName_tree == '' and ApplicationName_tree == '' and \
                   VariantName_tree == '':
                    print ' While Vehicle features '   
                    time.sleep(1)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Vehicle is - %s',Vehicle_Id)
                    
                    continue
               
                elif TestCaseName_tree == ''  and ApplicationName_tree == ''and\
                     VariantName_tree!= '' :
                    print 'Var_Val', Var_Val
                    
                    Write_Var = Variant_Value[Var_Val]
                    Variant_write(myAppl,Write_Var)                                         
                    Var_Val = Var_Val + 1
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Variant is - %s', VariantName_tree) 
                    time.sleep(3)
                        
                    continue
                
                elif TestCaseName_tree == ''  and ApplicationName_tree != '' and VariantName_tree!= '':
                
                    print ' This application started '  + ApplicationName_tree
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Application is - %s', ApplicationName_tree)
                    Sub_VariantName_tree = data[2]
                    print "Sub_VariantName_tree",Sub_VariantName_tree

                    
                    print  'dest_folder' + dest_folder
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    if not '_' in Sub_VariantName_tree :
                        
                        Sub_VariantName_tree = '_' + Sub_VariantName_tree

                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break
                    dest_folder = dest_folder +'\\' +  "01_Failsafe" ## folder_name_app         
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    for directories_d in dird:
                        if ApplicationName_tree == directories_d[3:]:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "ApplicationName_tree dest_folder",dest_folder
                            break                        

                
                    Category_Failsafe_Result_Folder = dest_folder
                          

                    Category_Master_Result_Report_Path=Master_Result_Report_Path+ "\\"+ ApplicationName_tree
                    
                    Report_dest_folder = Category_Failsafe_Result_Folder
                    distutils.dir_util.copy_tree(Category_Master_Result_Report_Path,Report_dest_folder)                                     
                    rep_name = ApplicationName_tree + '.xls'
                    Rep_Name = ApplicationName_tree  + str(Sub_VariantName_tree) + '.xls'
                    os.rename(os.path.join(Report_dest_folder,rep_name),os.path.join(Report_dest_folder,Rep_Name))
                    Rep_Path = os.path.join(Report_dest_folder,Rep_Name)
                    print "ApplicationName_tree",ApplicationName_tree
                    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet 

                    if os.path.exists(str(interface_sheet_path)):
                        xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)                    
                        xlapp.Application.Run("Interface_VBA.xls!module5.Hide_All",Rep_Path)
                        xlapp.Application.Run("Interface_VBA.xls!module14.Copysheet",ApplicationName_tree)
                    xlapp.Workbooks.Close()
                    if ApplicationName_tree!= 'CAN':
                        Actual_Test_Count=0
                        Executed_Test_Count=0
                        for key in Failsafe_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][ApplicationName_tree].iteritems():

                            Actual_Test_Count= Actual_Test_Count+1                        
                            print "Actual_Test_Count",Actual_Test_Count
                    Failsafe_Application_sheet = Book_Master_Test_sheet.sheet_by_name(ApplicationName_tree)
                    Failsafe_Application_sheet_End_row = Failsafe_Application_sheet.nrows
                    Failsafe_Application_sheet_End_col = Failsafe_Application_sheet.ncols 
               

                    time.sleep(1)

                    continue


                elif CANID_tree=='' and ECU_ID_tree=='' and CAN_Category_tree=='' and TestCaseName_tree == CAN_Test_02_str  and ApplicationName_tree != '' and VariantName_tree!= '':
                    
                    continue

                elif CANID_tree=='' and ECU_ID_tree=='' and CAN_Category_tree=='' and TestCaseName_tree == CAN_Test_03_str  and ApplicationName_tree != '' and VariantName_tree!= '':
                   
                    continue


                elif CANID_tree=='' and ECU_ID_tree=='' and CAN_Category_tree!='' and TestCaseName_tree == CAN_Test_02_str  and ApplicationName_tree != '' and VariantName_tree!= '':
                    
                    continue


                elif CANID_tree==''and ECU_ID_tree== '' and CAN_Category_tree!=''and TestCaseName_tree == CAN_Test_03_str  and ApplicationName_tree != '' and VariantName_tree!= '':
               
                    continue



                elif CANID_tree=='' and ECU_ID_tree!='' and CAN_Category_tree!='' and TestCaseName_tree == CAN_Test_02_str  and ApplicationName_tree != '' and VariantName_tree!= '':
                   
                    continue


                elif CANID_tree==''and ECU_ID_tree!= '' and CAN_Category_tree!=''and TestCaseName_tree == CAN_Test_03_str  and ApplicationName_tree != '' and VariantName_tree!= '':
                  
                    continue


                elif CANID_tree!=''and ECU_ID_tree!= '' and CAN_Category_tree!=''and TestCaseName_tree == CAN_Test_02_str  and ApplicationName_tree != '' and VariantName_tree!= '':


                    FAILSAFE_progressbar["maximum"] = 5
                    Failsafe_progress=0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress
                    
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]

                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                
                            break
                    dest_folder = dest_folder +'\\' + "01_Failsafe" ##folder_name_app       
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    for directories_d in dird:
                        if ApplicationName_tree == directories_d[3:]:
                            dest_folder = dest_folder + '\\' + directories_d
                            ssdest_folder = dest_folder
                            break

                    dest_folder = dest_folder + '\\' + CAN_Category_tree+ "\\" + ECU_ID_tree

                    try:
                        os.mkdir(dest_folder)
                    except :

                        pass
               
                    
                        
                    dest_folder=dest_folder+ "\\" +CANID_tree
                    try:
                        os.mkdir(dest_folder)
                    except:
                        pass

                    SigInfo=[]
                    SigInfo_temp=[]
                    SigNames=[]
                    SignalData=[]
                    SignalData1_temp=[]
                    sig_info={}
        
                                              

                
              
                    dest_Failsafe_Result_Folder=dest_folder
                    distutils.dir_util.copy_tree(Master_CANape_Failsafe_Path,dest_Failsafe_Result_Folder)
                    screenshot_Failsafe_Result_Folder = dest_Failsafe_Result_Folder + "\\"+"00_Screenshot"
                    Screenshot_Name= ECU_ID_tree + "_" + CANID_tree

                   
                    
                    

                   
                    for x in range(TestCase_Start_Row + 1,TestCase_End_Row):
                        SigInfo,SignalData1=Make_SignalData_Array(Primary_Test_sheet,x)

                    for k in range(0, Failsafe_Application_sheet_End_row):
                        
                        if Failsafe_Application_sheet.cell(k,9).value== CANID_tree :                                    
                            ApplicationSheet_TestCase_Start_Row= k
                  
                            for l in range(k, Failsafe_Application_sheet_End_row):
                                if Failsafe_Application_sheet.cell(l,0).value== "end_test_case" :
                                    ApplicationSheet_TestCase_End_Row =l                                           
                                    break
                          
                            Procedure_Test_sheet=Failsafe_Application_sheet
                            x=ApplicationSheet_TestCase_Start_Row+1
                            Make_SignalData_Array(Procedure_Test_sheet,x)                                   
                            x=ApplicationSheet_TestCase_End_Row-1
                            Make_SignalData_Array(Procedure_Test_sheet,x)

                    INPUT_SIGNAL_1_str= "INPUT_SIGNAL_1"

                    INPUT_SIGNAL_1_column	= 10
                    INPUT_SIGNAL_2_column   = 11

                    try:
                        for i in range (0,  Failsafe_Application_sheet_End_col):             # Loop for traversing through the EXCEL sheet
                            if Failsafe_Application_sheet.cell(0,i).value == INPUT_SIGNAL_1_str:         #This finds row and column of the particular variant
                                INPUT_SIGNAL_1_column =  i
                                INPUT_SIGNAL_2_column = INPUT_SIGNAL_1_column+1
                              
                    except:
                        INPUT_SIGNAL_1_column=10

                        INPUT_SIGNAL_2_column=11


                    try:
                        CAN_Value_1=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_1_column).value  
                    except:
                        CAN_Value_1="NA"
                    try:
                        CAN_Value_2=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_2_column).value
                    except:
                        CAN_Value_2="NA"

                    CAN_Message_Replacement(CAN_Value_1,CAN_Value_2)
                    Copy_TestCase_Data_to_PyMScript(Write_Var,ApplicationName_tree,CAN_Category_tree,ECU_ID_tree,CANID_tree,dest_Failsafe_Result_Folder)                

                    book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)                                   
                    JudgementsheetName = book1.sheet_by_name('Py-MScript')
                    if JudgementsheetName.cell(8,1).value == "NA":                                                      #check in judgementsheet for "A" or  "NA" for testcase applicable. if testcase not applicable it will disply TBD on GUI
                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.info('Test case skipped due to test case not present in judgementsheet %s')
                        Failsafe_Tree.item(curItem, text = CANID_tree, values ="TBD" )
                        continue

                    Failsafe_progress=Failsafe_progress+0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress

                    #Copy_TestCase_Data_to_PyMScript(Write_Var,ApplicationName_tree,CAN_Category_tree,ECU_ID_tree,CANID_tree,dest_Failsafe_Result_Folder)

                    Failsafe_progress=Failsafe_progress+0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     
                                
                    Execute_TestCase(dest_Failsafe_Result_Folder,SignalData1,SigInfo,sig_data_sheet,Screenshot_Name,curItem,CANID_tree,Meter_Navi_Layout_Name,Diag_Layout_Name,CAN_Category_tree)

                    Failsafe_progress=Failsafe_progress+2.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     
                  
                  
                    Reset_functionality(TestCase_Start_Row,TestCase_End_Row,Primary_Test_sheet,ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row,Failsafe_Application_sheet,myAppl)

                    Failsafe_progress=Failsafe_progress+1
                    FAILSAFE_progressbar["value"] = Failsafe_progress 




                    
                    continue

                elif CANID_tree!=''and ECU_ID_tree!= '' and CAN_Category_tree!=''and TestCaseName_tree == CAN_Test_03_str  and ApplicationName_tree != '' and VariantName_tree!= '':
     
                    FAILSAFE_progressbar["maximum"] = 5
                    Failsafe_progress=0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress  

                 

                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]

                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break
                    dest_folder = dest_folder +'\\' + "01_Failsafe" ##folder_name_app       
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    for directories_d in dird:
                        if ApplicationName_tree == directories_d[3:]:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "ApplicationName_tree dest_folder",dest_folder
                            ssdest_folder = dest_folder
                            break

                    
                    dest_folder = dest_folder + '\\' + CAN_Category_tree+ "\\" + ECU_ID_tree

                    try:
                        os.mkdir(dest_folder)
                    except Exception, e:
                        if e.errno != errno.EEXIST:
                            raise
                        pass                
        
                    dest_folder=dest_folder+ "\\" +CANID_tree
                    try:
                        os.mkdir(dest_folder)
                    except:
                        pass

                    if CAN_Category_tree=='Checksum':
                        CAN_Set_String='CHKKSM_set'
                        CAN_Value_String='CHKKSM_Value'
                    else:
                        CAN_Set_String='MSGGCNTR_Set'
                        CAN_Value_String='MSGGCNTR_Value'                   
                        

                    SigInfo=[]
                    SigInfo_temp=[]
                    SigNames=[]
                    SignalData=[]
                    SignalData1_temp=[]
                    sig_info={}

                    dest_Failsafe_Result_Folder=dest_folder
                    distutils.dir_util.copy_tree(Master_CANape_Failsafe_Path,dest_Failsafe_Result_Folder)
                    screenshot_Failsafe_Result_Folder = dest_Failsafe_Result_Folder + "\\"+"00_Screenshot"
                    Screenshot_Name= ECU_ID_tree + "_" + CANID_tree

                    

                        
                    for x in range(TestCase_Start_Row + 1,TestCase_End_Row):
                        SigInfo,SignalData1=Make_SignalData_Array(Primary_Test_sheet,x)


                    for k in range(0, Failsafe_Application_sheet_End_row):
                        
                        if Failsafe_Application_sheet.cell(k,9).value== CANID_tree :
                        

                              
                                ApplicationSheet_TestCase_Start_Row= k
                            
                                for l in range(k, Failsafe_Application_sheet_End_row):
                                    if Failsafe_Application_sheet.cell(l,0).value== "end_test_case" :
                                        ApplicationSheet_TestCase_End_Row =l
                                    
                                        break

               
                    for i in range(ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row):
                        if Failsafe_Application_sheet.cell(i,0).value== CAN_Value_String:
                            Procedure_Test_sheet=Failsafe_Application_sheet
                            x=i
                    
                            Make_SignalData_Array(Procedure_Test_sheet,x)

                        if Failsafe_Application_sheet.cell(i,0).value== CAN_Set_String:
                            Procedure_Test_sheet=Failsafe_Application_sheet
                            x=i
            
                            Make_SignalData_Array(Procedure_Test_sheet,x)

                    INPUT_SIGNAL_1_str= "INPUT_SIGNAL_1"

                    INPUT_SIGNAL_1_column	= 10
                    INPUT_SIGNAL_2_column   = 11

                    try:
                        for i in range (0,  Failsafe_Application_sheet_End_col):             # Loop for traversing through the EXCEL sheet
                            if Failsafe_Application_sheet.cell(0,i).value == INPUT_SIGNAL_1_str:         #This finds row and column of the particular variant
                                INPUT_SIGNAL_1_column =  i
                                INPUT_SIGNAL_2_column = INPUT_SIGNAL_1_column+1

                                
                    except:
                        INPUT_SIGNAL_1_column=10

                        INPUT_SIGNAL_2_column=11

                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.exception('Could not find string INPUT_SIGNAL_1 in Test Pattern.By default column 10 and 11 selected')

                    try:
                        CAN_Value_1=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_1_column).value  
                    except:
                        CAN_Value_1="NA"
                    try:
                        CAN_Value_2=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_2_column).value
                    except:
                        CAN_Value_2="NA"
                        

                    CAN_Message_Replacement(CAN_Value_1,CAN_Value_2)  
                    Copy_TestCase_Data_to_PyMScript(Write_Var,ApplicationName_tree,CAN_Category_tree,ECU_ID_tree,CANID_tree,dest_Failsafe_Result_Folder)

                    book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)                                   
                    JudgementsheetName = book1.sheet_by_name('Py-MScript')
                    if JudgementsheetName.cell(8,1).value == "NA":                                                      #check in judgementsheet for "A" or  "NA" for testcase applicable. if testcase not applicable it will disply TBD on GUI
                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.info('Test case skipped due to test case not present in judgementsheet %s')
                        Failsafe_Tree.item(curItem, text = CANID_tree, values ="TBD" )
                        continue


                    Failsafe_progress=Failsafe_progress+0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                         


                    x=ApplicationSheet_TestCase_End_Row-1
                    Make_SignalData_Array(Procedure_Test_sheet,x)

                    Failsafe_progress=Failsafe_progress+0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     
                                        
              
                    Execute_TestCase(dest_Failsafe_Result_Folder,SignalData1,SigInfo,sig_data_sheet,Screenshot_Name,curItem,CANID_tree,Meter_Navi_Layout_Name,Diag_layout_Name,CAN_Category_tree)

                    Failsafe_progress=Failsafe_progress+2.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     
                    Reset_functionality(TestCase_Start_Row,TestCase_End_Row,Primary_Test_sheet,ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row,Failsafe_Application_sheet,myAppl)

                    Failsafe_progress=Failsafe_progress+1
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     


                
                                   
                
                else:
                    flag = 0

                    FAILSAFE_progressbar["maximum"] = 5
                    Failsafe_progress=0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     
                    print "TestCaseName_tree",TestCaseName_tree
                    Sub_TestCaseName_tree = TestCaseName_tree[-2:]
                    ssdest_folder = ''
                    print "Sub_TestCaseName_tree",Sub_TestCaseName_tree
                    Screenshot_Name=Sub_TestCaseName_tree + "_" + ApplicationName_tree
                    print "Screenshot_Name",Screenshot_Name 
                    

                    ECU= "NA"
                    CANID="NA"
                    print "ApplicationName_tree",ApplicationName_tree
                    

                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]

                    for directories_d in dird:
                        if Sub_VariantName_tree in directories_d:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "Sub_VariantName_tree dest_folder",dest_folder
                            break
                    dest_folder = dest_folder +'\\' + "01_Failsafe" ##folder_name_app       
                    dird = [d for d in os.listdir(dest_folder) if os.path.isdir(os.path.join(dest_folder, d))]
                    for directories_d in dird:
                        if ApplicationName_tree == directories_d[3:]:
                            dest_folder = dest_folder + '\\' + directories_d
                            print "ApplicationName_tree dest_folder",dest_folder
                            ssdest_folder = dest_folder
                            break
                    dest_folder = dest_folder + '\\' + TestCaseName_tree
             
                    try:
                        os.mkdir(dest_folder)
                    except:   
                        pass
                    dest_Failsafe_Result_Folder=dest_folder

                    
                    if (flag == 1):
                        print "flag in if ",flag
                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.info("Test case not found in Master Configuration %s" + TestCaseName_tree + "  %s" + Sub_TestCaseName_tree)
                    else:
                     
                        distutils.dir_util.copy_tree(Master_CANape_Failsafe_Path,dest_folder)

                   

                    Application_start_test_str='start_'+ str(TestCaseName_tree)
                    Application_end_test_str='end_'+ str(TestCaseName_tree)

                    for k in range(0, Failsafe_Application_sheet_End_row):
                        if Failsafe_Application_sheet.cell(k,0).value== Application_start_test_str :
                            ApplicationSheet_TestCase_Start_Row= k
                            break

                    
                    for k in range(0, Failsafe_Application_sheet_End_row):
                        if Failsafe_Application_sheet.cell(k,0).value== Application_end_test_str:
                            ApplicationSheet_TestCase_End_Row= k
                            break    


                    INPUT_SIGNAL_1_str= "INPUT_SIGNAL_1"

                    INPUT_SIGNAL_1_column	= 10
                    INPUT_SIGNAL_2_column   = 11

                    try:
                        for i in range (0,  Failsafe_Application_sheet_End_col):             # Loop for traversing through the EXCEL sheet
                            if Failsafe_Application_sheet.cell(0,i).value == INPUT_SIGNAL_1_str:         #This finds row and column of the particular variant
                                INPUT_SIGNAL_1_column =  i
                                INPUT_SIGNAL_2_column = INPUT_SIGNAL_1_column+1

                               
                    except:
                        INPUT_SIGNAL_1_column=10

                        INPUT_SIGNAL_2_column=11

                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.exception('Could not find string INPUT_SIGNAL_1 in Test Pattern.By default column 10 and 11 selected')

                    try:
                        CAN_Value_1=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_1_column).value  
                    except:
                        CAN_Value_1="NA"
                    try:
                        CAN_Value_2=Failsafe_Application_sheet.cell(ApplicationSheet_TestCase_Start_Row+1,INPUT_SIGNAL_2_column).value
                    except:
                        CAN_Value_2="NA"

                    CAN_Message_Replacement(CAN_Value_1,CAN_Value_2)   
                    
              

                    Copy_TestCase_Data_to_PyMScript(Write_Var,ApplicationName_tree,Sub_TestCaseName_tree,ECU,TestCaseName_tree,dest_Failsafe_Result_Folder)

                    book1 = xlrd.open_workbook(Book_Master_Judgement_sheet,formatting_info=True)                                   
                    JudgementsheetName = book1.sheet_by_name('Py-MScript')
                    if JudgementsheetName.cell(9,1).value == "NA":                                                      #check in judgementsheet for "A" or  "NA" for testcase applicable. if testcase not applicable it will disply TBD on GUI
                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                        logging.info('Test case skipped due to test case not present in judgementsheet %s')
                        Failsafe_Tree.item(curItem, text = TestCaseName_tree, values ="TBD" )
                        continue

                    Failsafe_progress=Failsafe_progress+0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress 
                   
                    SigInfo=[]
                    SigInfo_temp=[]
                    SigNames=[]
                    SignalData=[]
                    SignalData1_temp=[]
                    sig_info={}


                    for x in range(TestCase_Start_Row + 1,TestCase_End_Row):
                        SigInfo,SignalData1=Make_SignalData_Array(Primary_Test_sheet,x)


                    for x in range(ApplicationSheet_TestCase_Start_Row+1,ApplicationSheet_TestCase_End_Row):                                
                        Procedure_Test_sheet=Failsafe_Application_sheet                                                            
                        SigInfo,SignalData1=Make_SignalData_Array(Procedure_Test_sheet,x)
                        

                    Failsafe_progress=Failsafe_progress+0.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress
                    
                    Execute_TestCase(dest_Failsafe_Result_Folder,SignalData1,SigInfo,sig_data_sheet,Screenshot_Name,curItem,TestCaseName_tree,Meter_Navi_Layout_Name,Diag_layout_Name,Sub_TestCaseName_tree)

                    Failsafe_progress=Failsafe_progress+2.5
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     

                    
                    Reset_functionality(TestCase_Start_Row,TestCase_End_Row,Primary_Test_sheet,ApplicationSheet_TestCase_Start_Row,ApplicationSheet_TestCase_End_Row,Failsafe_Application_sheet,myAppl)

                    Failsafe_progress=Failsafe_progress+1
                    FAILSAFE_progressbar["value"] = Failsafe_progress                     

                    Executed_Test_Count =Executed_Test_Count+1
                    print "Executed_Test_Count",Executed_Test_Count
                    print "Actual_Test_Count",Actual_Test_Count

                    if Executed_Test_Count == Actual_Test_Count:
                        xlapp1 = win32com.client.Dispatch("Excel.Application")
                        if os.path.exists(str(interface_sheet_path)):
                            xlapp1.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                            xlapp1.Application.Run("Interface_VBA.xls!module15.Hide_All")
                        xlapp1.Workbooks.Close()
                        #xlapp = win32com.client.dynamic.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet  
                    

                
               




            book3 = copy(book1)
            JudgementsheetName = book3.get_sheet(sheetNumber)
            print "Application_start_Row",Application_start_Row,Application_end_Row
            for i in range (Application_start_Row+30,  Application_end_Row+1+30):
                JudgementsheetName.write(i, 0, '')
                JudgementsheetName.write(i, 1, '')

                book3.save(Book_Master_Judgement_sheet)


            FAILSAFE_overall_progressbar["value"] = uid                
        except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.exception('Test case execution stopped abrubtly')                          

#**********************************************************************************************************************************************#
# Function Name: FailSafe_Testing
# Purpose : Tree Population, Data Extraction from Dispatch Sheet, Failsafe Conditions
# Inputs : NONE
# Outputs : NONE
#**********************************************************************************************************************************************#
    def FailSafe_Testing() : 
        #******************************Declaration*******************************************************#
        print "Failsafe Disctionary Making Start"
        global FailSafe_Dict,overall_progressbar_value,FailSafe_Display_Tree,curItem_FailSafe,A2l_Path
        global ECU_sensor_Dict,FailSafe_Result_Report_Destination,Output_Signal_Count,Expt_SetTime_Cols,Expt_Value_Cols,Actual_SetTime_Cols, \
               Actual_set_value_Cols,Test_result_varaint_Cols,temp_Master_Result_Report_FailSafe,Expected_DTC_Cols,Actual_DTC_Cols,DIMPSheet_Failsafe_CANID_List,myAppl, \
               same_signal_check,varaint_path,Failsafe_Delete_row,interface_sheet_path,Another_Signal_Name,DTC_string_path,DTC_string_path_1,Test_Sheet_Path
    
        Output_Signal_Count = 0
        count_hyperlinks = 1
        count_hyperlinks_CA = 1
        print count_hyperlinks
        temp_Master_Result_Report_FailSafe = ""
        varaint_path = 'Model Root/Driver Block/CANdb set/DIAG/ID7C3_TX/DIAG_CMD_NO/Value'
        FailSafe_Result_Report_Destination = []                
        FailSafe_Dict= OrderedDict()
        CanID_Data_FailSafe_Dict=OrderedDict()
        FailSafe_Dict_Tree= OrderedDict()       
        CanID_Data_FailSafe_Dict[VehicleName]= OrderedDict()
        FailSafe_Dict[VehicleName]= OrderedDict()
        FailSafe_Result_Folder=[] 
        ECU_Sensor_Array_dist_list = []
        CAN_ID_Array_list = []
        COUNT_YES = []
        Count_Ecu = []
        Expected_DTC_Cols = []
        Actual_DTC_Cols = []
        Expt_SetTime_Cols = []
        Expt_Value_Cols = []
        Test_result_varaint_Cols = []
        Actual_SetTime_Cols = []
        Actual_set_value_Cols = []
        Multiple_Input_signal_list =[]
        Multiple_Input_signal_After_list =[]
        Check_Multiple_Input_Signal = []
        Check_Multiple_Input_Signal_After = []
        #***************************************************************************************************#
        #*************************Sheet name path in python*************************************************#
        #DIMPSheet_Failsafe_CANID_List: The Failsafe dispatch sheet
        #DIMPSheet_Failsafe_Failsafe_List: The Failsafe signal Dispatch sheet
        #Master_Result_Report_FailSafe_WorkBook: The result workbook
        #Test_Result_Sheet_Failsafe:The result sheet in result workbook
        #Canid_Master_CANape_FailSafe_Path:The master canape configuration folder
        #dest_FailSafe_Result_Folder:The canape create for each ID in result folder
        #***************************************************************************************************#
        #*************************Lists and dictionary used*************************************************#
        #Inp_Sys_Var:value path of each signal

        #***************************************************************************************************#        
        print A2l_Path
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Failsafe Excution Started')                                                                 # Logging info in the log file
        # and ("/" not in Content)
        A2L_File_Open = open(A2l_Path,'r')          
        while True:
            Content = A2L_File_Open.readline()
            if (Content == "") :
                break
            elif (("aFS_Status" in Content)and ("/" not in Content)):
                Output_Signal_Count = Output_Signal_Count + 1
        Output_Signal_Count = 5
##        with open(A2l_Path) as A2L_File_Open:
##            Content = A2L_File_Open.readlines()
##            print Content
##            if (("aFS_Status" in Content)):
##                print "Match"
##                Output_Signal_Count = Output_Signal_Count + 1
##        Output_Signal_Count = Output_Signal_Count - 1 

##        shutil.copy(Master_Result_Report_FailSafe,VehicleNameFolder[0])
##        temp_Master_Result_Report_FailSafe = VehicleNameFolder[0] + "\\" + "Master_Result_Report_FailSafe.xls"
##        print temp_Master_Result_Report_FailSafe 
##        Master_Result_Report_FailSafe_WorkBook = xlrd.open_workbook(str(temp_Master_Result_Report_FailSafe),formatting_info=True) #This opens Result DISPATCH SHEET Workbook
##        Test_Result_Sheet_Failsafe = Master_Result_Report_FailSafe_WorkBook.sheet_by_index(1)   #The Failsafe Sheet in DISPATCH SHEET WORKBOOK
##        Test_Result_Sheet_Failsafe_Col = Test_Result_Sheet_Failsafe.ncols
##        Test_Result_Sheet_Failsafe_Row = Test_Result_Sheet_Failsafe.nrows
##
##        for i in range (0, Test_Result_Sheet_Failsafe_Row):
##            for j in range (0,Test_Result_Sheet_Failsafe_Col) :
##                if Test_Result_Sheet_Failsafe.cell(i,j).value == "Expected Set Time" :                    
##                    Expt_SetTime_Cols.append(j)
##                if Test_Result_Sheet_Failsafe.cell(i,j).value == "Actual Set Time":
##                    Actual_SetTime_Cols.append(j)
##                if Test_Result_Sheet_Failsafe.cell(i,j).value == "Actual Set Value":
##                    Actual_set_value_Cols.append(j)
##                if Test_Result_Sheet_Failsafe.cell(i,j).value == "Expected Value":
##                    Expt_Value_Cols.append(j)
##                if Test_Result_Sheet_Failsafe.cell(i,j).value ==  "Test Result":
##                    Test_result_varaint_Cols.append(j)
##
##        
##        FAILSAFE_progressbar["maximum"] = 14
##        excel= win32com.client.dynamic.Dispatch("Excel.Application")
##        workbook = excel.Workbooks.Open(str(temp_Master_Result_Report_FailSafe))
##        Test_Result_Sheet_Failsafe = workbook.Sheets(2)
##        Deleted_row_count = 0
##        for i in range (1, Test_Result_Sheet_Failsafe_Row):
##            cell_value_test = Test_Result_Sheet_Failsafe.Cells(i,2).value
##            if (cell_value_test != "JT1" and cell_value_test != "JT2" and len(str(cell_value_test)) == 3 ) :
##                id_found_test_result = False
##                id_found_test_result = 0
##                for a in range (0 , len(CAN_ID_Array_list)):
##                    if cell_value_test == CAN_ID_Array_list[a]:
##                        id_found_test_result = True
##
##                if id_found_test_result == False:
##                    Test_Result_Sheet_Failsafe.Rows(i).Entirerow.Hidden = True
####                    Test_Result_Sheet_Failsafe.Rows(i).Delete()
##                    i = i - 1
##                    Deleted_row_count = Deleted_row_count + 1
##                        
##        workbook.save
##        workbook.close
##        print "Number of Row Deleted is " , Deleted_row_count
##        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
##            format='%(asctime)s - %(levelname)s - %(message)s')             
##        logging.info('Master Result Report Updated With CanId ')                                                                 # Logging info in the log file



                
        try:

            for k in range (0, len(Variant)):
                counter = 0
                Region_variant_number = RegionName  + '_' + str(int(Variant_Value[k]))
                FailSafe_Result_Folder.append(VehicleNameFolder[k]+"\\09_FailSafe")
                FailSafe_Result_Report_Destination.append(VehicleNameFolder[k]+"\\09_FailSafe")
                print Master_Result_FailSafe_Path,'FailSafe_Result_Report_Destination',FailSafe_Result_Report_Destination
                copy_folder(Master_Result_FailSafe_Path,FailSafe_Result_Report_Destination)
                FailSafe_Dict[VehicleName][Region_variant_number]  = OrderedDict()
                Var_Row = 0     #Store row number of Variant in DISPATCH SHEET 
                Var_col = 0     #Store column number of Variant in DISPATCH SHEET
                CANID_row=0     #Store row number of the CANID in DISPATCH SHEET 
                CANID_col=0     #Store column number of the CANID in DISPATCH SHEET 
                ECU_Sensor_col=0
                ECU_Sensor_Row=0
                ECU_Sensor_Array_list = ['ALL']            
                DIMPSheet_Failsafe_CANID_List = WorkBook.sheet_by_name("Failsafe_Input_Test_Pattern")
                DIMPSheet_Failsafe_Failsafe_List = WorkBook.sheet_by_name("Failsafe_signals_test_pattern")                  
                DIMPSheet_Failsafe_CANID_List_Col = DIMPSheet_Failsafe_CANID_List.ncols     
                DIMPSheet_Failsafe_CANID_List_Row = DIMPSheet_Failsafe_CANID_List.nrows 
                DIMPSheet_Failsafe_Failsafe_List_Col = DIMPSheet_Failsafe_Failsafe_List.ncols    
                DIMPSheet_Failsafe_Failsafe_List_Row = DIMPSheet_Failsafe_Failsafe_List.nrows   
                print  'Row and Column'
                print  DIMPSheet_Failsafe_Failsafe_List_Col
                print  DIMPSheet_Failsafe_Failsafe_List_Row
                Var_str =  'Variant_' +  str(int(Variant_Value[k]))
                print Var_str,'Var_str'
                #*********************************************To find the necessary Details in Failsafe Dispatch sheet and Store in Dictionary**************************************#
                for i in range (0,  DIMPSheet_Failsafe_CANID_List_Row):             # Loop for traversing through the EXCEL sheet
                    for j in range(0,DIMPSheet_Failsafe_CANID_List_Col):
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == 'ECU/Sensor':         #This finds row and column of the particular variant
                            ECU_Sensor_col =  j
                            ECU_Sensor_Row = i
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == Var_str :         #This finds row and column of the particular variant
                            Var_col =  j
                            Var_Row = i
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == "JT1" :
                            JT1_time_col = j
                            JT1_time_row = i+1
                            JT1_value_col = j+1
                            JT1_value_row = i+1
                            JT1_DTC_col = j+2
                            JT1_DTC_row = i+1
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == "JT2" :
                            JT2_time_col = j
                            JT2_time_row = i+1
                            JT2_value_col = j+1
                            JT2_value_row = i+1
                            JT2_DTC_col = j+2
                            JT2_DTC_row = i+1
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == "Message-Counter" :
                            message_counter_time_col = j
                            message_counter_time_row = i+1
                            message_counter_value_col = j+1
                            message_counter_value_row = i+1
                            message_counter_DTC_col = j+2
                            message_counter_DTC_row = i+1
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == "Checksum" :
                            checksum_time_col = j
                            checksum_time_row = i+1
                            checksum_value_col = j+1
                            checksum_value_row = i+1
                            checksum_DTC_col = j+2
                            checksum_DTC_row = i+1
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == "Voltage_Check" :
                            voltage_check_time_col = j
                            voltage_check_time_row = i+1                        
                            voltage_check_value_col = j+1
                            voltage_check_value_row = i+1
                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == "Failsafe_Signals" :
                            Failsafe_Signals_col = j
                            Failsafe_Signals_row = i


                        if DIMPSheet_Failsafe_CANID_List.cell(i,j).value == 'CAN ID':        #This finds row and column of CANIDs in MSG_COUNTER_DETAIL
                             CANID_col =  j
                             CANID_row = i
                        else:
                            continue

                print "Failsafe_Signals_col"
                print Failsafe_Signals_col
##                CanID_Data_FailSafe_Dict[VehicleName] = OrderedDict()
                CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number] = OrderedDict()
               
                for i in range (Var_Row + 2, DIMPSheet_Failsafe_CANID_List_Row):
                    if (DIMPSheet_Failsafe_CANID_List.cell(i,Var_col).value=='Y'):
                        counter = counter +1
                        ECU_Sensor_value=DIMPSheet_Failsafe_CANID_List.cell(i,ECU_Sensor_col).value
                        Can_id_value = DIMPSheet_Failsafe_CANID_List.cell(i,CANID_col).value
                        Canid_Master_CANape_FailSafe_Path = Master_CANape_FailSafe_Path + "\\ID" + str(Can_id_value)
                        dest_FailSafe_Result_Folder = FailSafe_Result_Folder[k] + "\\ID" + str(Can_id_value)
                        #copy_folder(Canid_Master_CANape_FailSafe_Path,dest_FailSafe_Result_Folder)        # This function copies files and folders from MASTER CANAPE CONFIG
                        if ECU_Sensor_value in ECU_Sensor_Array_list:
                            FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value] = OrderedDict()
                            CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value] = OrderedDict()  
                            CAN_ID_Array_list.append(Can_id_value)
                        else :
                            CAN_ID_Array_list.append(Can_id_value)
                            ECU_Sensor_Array_list.append(ECU_Sensor_value)
                            FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value] = OrderedDict()
                            FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value] = OrderedDict()
                            CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value] = OrderedDict()
                            CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value] = OrderedDict()
                        
                        
##                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value] = OrderedDict()
##                        CanID_Data_FailSafe_Dict = FailSafe_Dict
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['JT1_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT1_time_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['JT1_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT1_value_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['JT1_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT1_DTC_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['JT2_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT2_time_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['JT2_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT2_value_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['JT2_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT2_DTC_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Message_Counter_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,message_counter_time_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Message_Counter_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,message_counter_value_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Message_Counter_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,message_counter_DTC_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Checksum_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,checksum_time_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Checksum_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,checksum_value_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Checksum_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,checksum_DTC_col).value                    
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Voltage_check_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,voltage_check_time_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Voltage_check_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,voltage_check_value_col).value
                        CanID_Data_FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]['Failsafe_Signals'] = DIMPSheet_Failsafe_CANID_List.cell(i,Failsafe_Signals_col).value
##                CanID_Data_FailSafe_Dict = copy.deepcopy(FailSafe_Dict)
##                for i in range (Var_Row + 2, DIMPSheet_Failsafe_CANID_List_Row):
##                    if (DIMPSheet_Failsafe_CANID_List.cell(i,Var_col).value=='Y'):
##                        ECU_Sensor_value=DIMPSheet_Failsafe_CANID_List.cell(i,ECU_Sensor_col).value
##                        print "ECU_Sensor_value",ECU_Sensor_value
##                        Can_id_value = DIMPSheet_Failsafe_CANID_List.cell(i,CANID_col).value
##                        for variant_name_key in FailSafe_Dict[VehicleName].iteritems():                            
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['JT1_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT1_time_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['JT1_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT1_value_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['JT1_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT1_DTC_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['JT2_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT2_time_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['JT2_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT2_value_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['JT2_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,JT2_DTC_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Message_Counter_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,message_counter_time_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Message_Counter_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,message_counter_value_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Message_Counter_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,message_counter_DTC_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Checksum_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,checksum_time_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Checksum_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,checksum_value_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Checksum_DTC'] = DIMPSheet_Failsafe_CANID_List.cell(i,checksum_DTC_col).value                    
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Voltage_check_time'] = DIMPSheet_Failsafe_CANID_List.cell(i,voltage_check_time_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Voltage_check_value'] = DIMPSheet_Failsafe_CANID_List.cell(i,voltage_check_value_col).value
##                            CanID_Data_FailSafe_Dict[VehicleName][variant_name_key[0]][ECU_Sensor_value][Can_id_value]['Failsafe_Signals'] = DIMPSheet_Failsafe_CANID_List.cell(i,Failsafe_Signals_col).value
                COUNT_YES.append(counter)
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                format='%(asctime)s - %(levelname)s - %(message)s')             
                logging.info('Dictionary Created: %s',CanID_Data_FailSafe_Dict)
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                format='%(asctime)s - %(levelname)s - %(message)s')             
                logging.info('Dictionary1 Created: %s',FailSafe_Dict)
                
                
            #**************************************************************************************************************************************************************************************************#
            
            FAILSAFE_progressbar["maximum"] = 14
            overall_progressbar_value = 0
            max_overall_progressbar_value  = 0
            for variant_name_key in FailSafe_Dict[VehicleName].keys():
                for ecusensort in FailSafe_Dict[VehicleName][variant_name_key].keys():
                    for canidc in FailSafe_Dict[VehicleName][variant_name_key][ecusensort].values():
                        max_overall_progressbar_value += 1
            
            FAILSAFE_overall_progressbar["maximum"] = max_overall_progressbar_value * 5
            FAILSAFE_progressbar["value"] = 1
            FAILSAFE_overall_progressbar["value"] = overall_progressbar_value        
            shutil.copy(Master_Result_Report_FailSafe,FailSafe_Result_Report_Destination[0])
            temp_Master_Result_Report_FailSafe = FailSafe_Result_Report_Destination[0] + "\\" + "Master_Result_Report_FailSafe.xls"
            print temp_Master_Result_Report_FailSafe 
            Master_Result_Report_FailSafe_WorkBook = xlrd.open_workbook(str(temp_Master_Result_Report_FailSafe),formatting_info=True) #This opens Result DISPATCH SHEET Workbook
            Test_Result_Sheet_Failsafe = Master_Result_Report_FailSafe_WorkBook.sheet_by_index(1)   #The Failsafe Sheet in DISPATCH SHEET WORKBOOK
            Test_Result_Sheet_Failsafe_Col = Test_Result_Sheet_Failsafe.ncols
            Test_Result_Sheet_Failsafe_Row = Test_Result_Sheet_Failsafe.nrows
            #***************************To find the necessary details in master result sheet and result folder *************************************************************************************************#
            for i in range (0, Test_Result_Sheet_Failsafe_Row):
                for j in range (0,Test_Result_Sheet_Failsafe_Col) :
                    if Test_Result_Sheet_Failsafe.cell(i,j).value == "Expected Set Time" :                    
                        Expt_SetTime_Cols.append(j)
                    if Test_Result_Sheet_Failsafe.cell(i,j).value == "Actual Set Time":
                        Actual_SetTime_Cols.append(j)
                    if Test_Result_Sheet_Failsafe.cell(i,j).value == "Actual Set Value":
                        Actual_set_value_Cols.append(j)
                    if Test_Result_Sheet_Failsafe.cell(i,j).value == "Expected Value":
                        Expt_Value_Cols.append(j)
                    if Test_Result_Sheet_Failsafe.cell(i,j).value ==  "Test Result":
                        Test_result_varaint_Cols.append(j)                        
                    if Test_Result_Sheet_Failsafe.cell(i,j).value ==  "Expected DTC":
                        Expected_DTC_Cols.append(j)
                    if Test_Result_Sheet_Failsafe.cell(i,j).value ==  "Actual Set DTC":
                        Actual_DTC_Cols.append(j)
            
            #******************************************************************************************************************************************************************************************************#
            excel= win32com.client.dynamic.Dispatch("Excel.Application")            
            delete_row_workbook = excel.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
            delete_row_worksheet = delete_row_workbook.Sheets(2)
            for a in range (0 , len(CAN_ID_Array_list)):
                delete_row_worksheet.Cells(a + 1 ,1).Value = CAN_ID_Array_list[a]
            excel.Application.Run("Interface_VBA.xls!module8.Delete_rows",temp_Master_Result_Report_FailSafe,len(CAN_ID_Array_list))
            ##excel.Application.Quit()
            delete_row_workbook.Save()
            delete_row_workbook.Close()
            excel= win32com.client.dynamic.Dispatch("Excel.Application")
            print temp_Master_Result_Report_FailSafe,'temp_Master_Result_Report_FailSafe'
            workbook = excel.Workbooks.Open(str(temp_Master_Result_Report_FailSafe))
            Test_Result_Sheet_Failsafe = workbook.Sheets(2)
            Deleted_row_count = 0
            Test_Result_Sheet_Failsafe.Cells(2,3).Value = VehicleName
            Test_Result_Sheet_Failsafe.Cells(3,3).Value = RegionName
            Test_Result_Sheet_Failsafe.Cells(4,3).Value = PartNo


   
##            for i in range (1, Test_Result_Sheet_Failsafe_Row):
##                cell_value_test = Test_Result_Sheet_Failsafe.Cells(i,2).Value
##                if (i == 100 ) :
##                    FAILSAFE_progressbar["value"] = 1.2
##                if (i == 200 ) :
##                    FAILSAFE_progressbar["value"] = 1.4
##                if (i == 300 ) :
##                    FAILSAFE_progressbar["value"] = 1.6
##                if (i == 400 ) :
##                    FAILSAFE_progressbar["value"] = 1.8
##                if (i == 500 ) :
##                    FAILSAFE_progressbar["value"] = 2
##                if (cell_value_test != "JT1" and cell_value_test != "JT2" and len(str(cell_value_test)) == 3 ) :
##                    id_found_test_result = False
##                    id_found_test_result = 0
##                    for a in range (0 , len(CAN_ID_Array_list)):
##                        if cell_value_test == CAN_ID_Array_list[a]:
##                            id_found_test_result = True
##                    if id_found_test_result == False:
##                        Test_Result_Sheet_Failsafe.Rows(i).EntireRow.Hidden = True
##                        #Test_Result_Sheet_Failsafe.Rows(i).Entirerow.Hidden = True
####                        Test_Result_Sheet_Failsafe.Rows(i).Delete()
##                        i = i - 1
##                        Deleted_row_count = Deleted_row_count + 1
                        
                            
            workbook.Save()
            workbook.Close()
            print "Number of Row Deleted is " , Deleted_row_count
 
            
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                format='%(asctime)s - %(levelname)s - %(message)s')             
            logging.info('Master Result Report Updated With CanId ')                                                                 # Logging info in the log file
            

            print CanID_Data_FailSafe_Dict
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                format='%(asctime)s - %(levelname)s - %(message)s')             
            logging.info('Dictionary Created')                                                                 # Logging info in the log file

                    
            
            ECU_sensor_drop['values'] = ECU_Sensor_Array_list                                                 #Creating dropdown option for ECU/sensor
            ECU_sensor_drop.set(ECU_Sensor_Array_list[0])        
            ECU_sensor_drop.pack(side=LEFT,padx=5,pady=5)        
            Failsafe_categary_list = ('ALL','Message_Counter','Checksum','JT1','JT2','Failsafe_Signals')                        # to store failsafe category lists
            failsafe_cat__drop['values'] = Failsafe_categary_list
            failsafe_cat__drop.set(Failsafe_categary_list[0])
            failsafe_cat__drop.pack(side=RIGHT,padx=5,pady=5)                                                            #Creating dropdown option for Failsafe Category
            uid_Failsafe_prev=uid
            FailSafe_Display_Tree = construct_JSON_tree(FailSafe_Dict,frame7)
            
            
            #Message_Counter_Tree = construct_JSON_tree(Message_Counter_Dict,frame8)    #This makes Message Counter Tree
            curItem_FailSafe= uid_Failsafe_prev+1     #Since uid doest become 0
            Failsafe_vehicle_id = FailSafe_Display_Tree.item(curItem_FailSafe, 'text')   #Extracts information from Message Counter Tree
            FAILSAFE_vehicle_id_entry["state"] = NORMAL
            FAILSAFE_vehicle_id_entry.insert(0, Failsafe_vehicle_id)            #Fills the space in Message Counter GUI with required information
            FAILSAFE_vehicle_id_entry["state"] = DISABLED
            FAILSAFE_CAT_entry["state"] = NORMAL
            FAILSAFE_CAT_entry.insert(0,failsafe_cat__drop.get())
            FAILSAFE_CAT_entry["state"] = DISABLED

            curItem_FailSafe=curItem_FailSafe+1   #This is used to jump over the items not required in Tree for displaying in GUI

            counter_row=0 #Simple counter variable for incrementing
            Failsafe_Variant_Increment = 0
            for variant_name_key in FailSafe_Dict[VehicleName].iteritems():
                count_hyperlinks = 0
##                count_hyperlinks_CA = 1
                FAILSAFE_progressbar["value"] = 3                
                FAILSAFE_variant_entry["state"] = NORMAL
                FAILSAFE_CAN_ID_entry["state"] = NORMAL
                FAILSAFE_ECU_entry["state"] = NORMAL
                FAILSAFE_result_entry["state"] = NORMAL
                FAILSAFE_variant_entry.delete(0,END)      #Clears the space in Fail safe GUI  
                FAILSAFE_CAN_ID_entry.delete(0,END)       #Clears the space in Fail safe GUI 
                FAILSAFE_result_entry.delete(0,END)       #Clears the space in Fail safe GUI
                FAILSAFE_ECU_entry.delete(0,END)
                FailSafe_Display_Tree.selection_set(curItem_FailSafe)   #Highlighting the item in the Tree
                Variant_Name = variant_name_key[0]            
                FAILSAFE_variant_entry["state"] = NORMAL            
                FAILSAFE_variant_entry.insert(0, Variant_Name)
                FAILSAFE_variant_entry["state"] = DISABLED
                print Variant_Value,'Variant_Value',Failsafe_Variant_Increment
                print Variant_Name
                Write_Var = Variant_Value[Failsafe_Variant_Increment]
                print Write_Var
                if Failsafe_Variant_Increment != 0:
                    curItem_FailSafe = curItem_FailSafe + 1
                    print 'VehicleNameFolder[Failsafe_Variant_Increment]',VehicleNameFolder[Failsafe_Variant_Increment]
                    temp_Master_Result_Report_FailSafe = VehicleNameFolder[Failsafe_Variant_Increment] + "\\09_FailSafe"
                    shutil.copy(Master_Result_Report_FailSafe,temp_Master_Result_Report_FailSafe)
                    temp_Master_Result_Report_FailSafe = VehicleNameFolder[Failsafe_Variant_Increment] + "\\09_FailSafe" + "\\" + "Master_Result_Report_FailSafe.xls"
                    excel= win32com.client.dynamic.Dispatch("Excel.Application")            
                    delete_row_workbook = excel.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    delete_row_worksheet = delete_row_workbook.Sheets(1)
                    for a in range (0 , len(CAN_ID_Array_list)):
                        delete_row_worksheet.Cells(a + 1 ,1).Value = CAN_ID_Array_list[a]
                    excel.Application.Run("Interface_VBA.xls!module8.Delete_rows",temp_Master_Result_Report_FailSafe,len(CAN_ID_Array_list))
                    ##excel.Application.Quit()
                    delete_row_workbook.Save()
                    delete_row_workbook.Close()
            
                    
                print Power_Supply_path
                myAppl.Variable(Power_Supply_path).Write(1)   # Switch on Power Supply to Write
                
                        ##Instrumentation().ActiveLayout.Normalize()
                time.sleep(0.5)
                if AdasECU == 'Dual ADAS':
                    
                    myAppl.Variable('Model Root/Driver Block/CANdb set/ADAS2/ADAS2_EXIST/Value').Write(0)  
                    
                    myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)         # Write Variant Code to the Control Desk (Config Layout) 
                    time.sleep(1)
                    myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C3_TX/DIAG_CMD_NO/Value').Write(4) 
                    time.sleep(.5)
                    myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C3_TX/DIAG_CMD_NO/Value').Write(0) 
                    
                else :
                    time.sleep(2)
                    myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)         # Write Variant change value to variable 
                    time.sleep(0.5)
                    myAppl.Variable(varaint_path).Write(4)      
                    time.sleep(0.5)
                    myAppl.Variable(varaint_path).Write(0)      
                    time.sleep(2)
                    
                myAppl.Variable(Power_Supply_path).Write(129)   #Reset power supply after writing Variant Code  
                time.sleep(0.5)
                myAppl.Variable(Power_Supply_path).Write(1) 
                time.sleep(0.5)
                myAppl.Variable(Power_Supply_path).Write(129)           
                #FailSafe_Dict[VehicleName][Region_variant_number][ECU_Sensor_value][Can_id_value]
                for ecusensort in FailSafe_Dict[VehicleName][Variant_Name].iteritems():
                    
                    print "ECU_Sensor_value"
                    FAILSAFE_ECU_entry["state"] = NORMAL
                    FAILSAFE_CAN_ID_entry["state"] = NORMAL
                    FAILSAFE_CAN_ID_entry.delete(0,END)
                    FAILSAFE_result_entry["state"] = NORMAL
                    FAILSAFE_result_entry.delete(0,END)                
                    FAILSAFE_ECU_entry.delete(0,END)                                
                    curItem_FailSafe = curItem_FailSafe + 1 
                    FailSafe_Display_Tree.selection_set(curItem_FailSafe)   #Highlighting the item in the Tree
                    Ecu_Name_curItem = FailSafe_Display_Tree.item(curItem_FailSafe, 'text')
                    FAILSAFE_ECU_entry["state"] = NORMAL
                    FAILSAFE_ECU_entry.insert(0, Ecu_Name_curItem)
                    print Ecu_Name_curItem
                    FAILSAFE_ECU_entry["state"] = DISABLED
                    for canidc in FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem].iteritems():
                        FAILSAFE_progressbar["value"] = 4
                        FAILSAFE_CAN_ID_entry["state"] = NORMAL
                        FAILSAFE_CAN_ID_entry.delete(0,END)
                        curItem_FailSafe=curItem_FailSafe+1
                        FailSafe_Display_Tree.selection_set(curItem_FailSafe)
                        CANID_Failsafe = FailSafe_Display_Tree.item(curItem_FailSafe, 'text')
                        FAILSAFE_CAN_ID_entry["state"] = NORMAL
                        FAILSAFE_CAN_ID_entry.insert(0, CANID_Failsafe )
                        print CANID_Failsafe
                        FAILSAFE_CAN_ID_entry["state"] = DISABLED
                        FAILSAFE_result_entry["state"] = NORMAL
                        FAILSAFE_result_entry.delete(0,END)
                        
                        Dest_Procedure_failsafe_result = FailSafe_Result_Folder[k] + "\\ID" + canidc[0]
                        print Dest_Procedure_failsafe_result
                        Test_Sheet = Test_Sheet_Path + '\\' + VehicleName + '\\' + 'Master_TestPattern_FLS.xls'
                        Open_Test_case_sheet = xlrd.open_workbook(Test_Sheet)                                                                       # Opening the test case sheet
                        FailSafe_Test_Procedure = "FailSafe_Test_Procedure"
                        Signal_data_sheet_name = sig_data_sheet_str
                        DTC_string_name = "DTC_String"
                        DTC_string_name_1 = "DTC_String_1"
                        FailSafe_Test_Procedure_sheet = Open_Test_case_sheet.sheet_by_name(FailSafe_Test_Procedure)
                        Signal_data_sheet = Open_Test_case_sheet.sheet_by_name(Signal_data_sheet_name)
                        message_Counter_Inp_Sys_Var_row = 0
                        message_Counter_Sys_Var_Set_row = 0
                        Checksum_Inp_Sys_Var_row = 0
                        DTC_string_path = 0
                        DTC_string_path_1 = 0
                        for row_signal_data in range(0,Signal_data_sheet.nrows):
                            if Signal_data_sheet.cell(row_signal_data,0).value == DTC_string_name:
                                DTC_string_path = Signal_data_sheet.cell(row_signal_data,1).value
                            if Signal_data_sheet.cell(row_signal_data,0).value == DTC_string_name_1:
                                DTC_string_path_1 = Signal_data_sheet.cell(row_signal_data,1).value 
                                
                        for row_FailSafe_Test_Procedure in range (0,FailSafe_Test_Procedure_sheet.nrows):
                            if FailSafe_Test_Procedure_sheet.cell(row_FailSafe_Test_Procedure,4).value == "ID" + CANID_Failsafe + "_Message_Counter_Set" :
                                message_Counter_Sys_Var_Set_row = row_FailSafe_Test_Procedure                            
                            if FailSafe_Test_Procedure_sheet.cell(row_FailSafe_Test_Procedure,4).value == "ID" + CANID_Failsafe + "_Message_Counter_Value" :
                                message_Counter_Inp_Sys_Var_row = row_FailSafe_Test_Procedure
                            if FailSafe_Test_Procedure_sheet.cell(row_FailSafe_Test_Procedure,4).value == "ID" + CANID_Failsafe + "_Checksum_Set" :
                                Checksum_Sys_Var_Set_row = row_FailSafe_Test_Procedure
                            if FailSafe_Test_Procedure_sheet.cell(row_FailSafe_Test_Procedure,4).value == "ID" + CANID_Failsafe + "_Checksum_Value" :
                                Checksum_Inp_Sys_Var_row = row_FailSafe_Test_Procedure
                            if FailSafe_Test_Procedure_sheet.cell(row_FailSafe_Test_Procedure,4).value == "ID" + CANID_Failsafe + "_CANON/OFF_Value" :
                                CANONOFF_Inp_Sys_Var_row = row_FailSafe_Test_Procedure

                        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                            format='%(asctime)s - %(levelname)s - %(message)s')             
                        logging.info('Failsafe Canid System Var Extracted')                                                                 # Logging info in the log file

                        Expected_Result = ['','','','']
                        Expected_DTC = ""
                        for failsafe_Categate_loop in range(0,len(Failsafe_categary_list)):
                            FAILSAFE_progressbar["value"] = 5
                            word = ''
                            FailSafe_Category_Pre = Failsafe_categary_list[failsafe_Categate_loop]
                            FailSafe_Category_Present = Failsafe_categary_list[failsafe_Categate_loop]
                            FAILSAFE_CAT_entry["state"] = NORMAL
                            FAILSAFE_CAT_entry.delete(0,END)
                            FAILSAFE_CAT_entry.insert(0,FailSafe_Category_Present)
                            FAILSAFE_CAT_entry["state"] = DISABLED
                            FAILSAFE_result_entry["state"] = NORMAL
                            FAILSAFE_variant_entry.delete(0,END)
                            FAILSAFE_result_entry.delete(0,END)
                            Message_counter_Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(message_Counter_Inp_Sys_Var_row,6).value
                            Checksum_Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(Checksum_Inp_Sys_Var_row,6).value
                            print "Debug",Message_counter_Inp_Sys_Var
                            print "Debug",Checksum_Inp_Sys_Var                             
                            print CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Failsafe_Signals']
                            check_dependent_signal = False
                            Result_Dependent_Signal_list = []
                            Result_Dependent_Signal_list_value = []
                            Result_Dependent_Signal_Comparison_list_value = []
                            Multiple_Input_signal_After_list = []
                            Multiple_Input_signal_list = []
                            Check_Multiple_Input_Signal = False
                            Check_Multiple_Input_Signal_After = False
                            
                            
                            if Failsafe_categary_list[failsafe_Categate_loop] == "Message_Counter" and CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Message_Counter_value']  != '':
                                Volt_Ref = 0
                                Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(message_Counter_Inp_Sys_Var_row,6).value
                                Sys_Var_Set = FailSafe_Test_Procedure_sheet.cell(message_Counter_Sys_Var_Set_row,6).value
                                print CANID_Failsafe
                                print Inp_Sys_Var
                                print Sys_Var_Set
                                Dpdt_Signal_Set = None
                                Dependent_signal = None
                                Input_Value = 1
                                Reset_Value = 0
                                Dpdt_Value = 0
                                number_of_Results = 1
                                Reset_Time = 0
                                Expected_Result[0] = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Message_Counter_value']
                                Expected_DTC = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Message_Counter_DTC']
                                Exec_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Message_Counter_time']
                                Exec_Type = "TYPE_A"
                                Canid_Master_CANape_FailSafe_Path = Master_CANape_FailSafe_Path + "\\ID" + CANID_Failsafe
                                dest_FailSafe_Result_Folder = FailSafe_Result_Folder[Failsafe_Variant_Increment] + "\\ID" + CANID_Failsafe + "\\" + FailSafe_Category_Pre
                                copy_folder(Canid_Master_CANape_FailSafe_Path,dest_FailSafe_Result_Folder)
                                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                                    format='%(asctime)s - %(levelname)s - %(message)s')             
                                logging.info('Failsafe Test Procedure For Message Counter Started')                                                                 # Logging info in the log file                            
                                count_hyperlinks,count_hyperlinks_CA = Fail_safe_Test_procedure(Inp_Sys_Var,Sys_Var_Set,Dependent_signal,Dpdt_Signal_Set,Exec_Type, \
                                                         Dest_Procedure_failsafe_result,FailSafe_Category_Pre,Dpdt_Value,Reset_Value,Input_Value, \
                                                         Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                                         Expected_Result,count_hyperlinks,Expected_DTC,FailSafe_Result_Folder[Failsafe_Variant_Increment],word,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                                         check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list, \
                                                         Open_Test_case_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,Multiple_Input_signal_After_list)
                                #make one while loop which monitor delay_variable to 4                            
                                        
                            elif Failsafe_categary_list[failsafe_Categate_loop] == "Checksum" and CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Checksum_value']  != '':
                                Volt_Ref = 0
                                Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(Checksum_Inp_Sys_Var_row,6).value   #value path
                                Sys_Var_Set = FailSafe_Test_Procedure_sheet.cell(Checksum_Sys_Var_Set_row,6).value   #set path
                                Dpdt_Signal_Set = None
                                Dependent_signal = None
                                Input_Value = 1
                                Reset_Value = 0
                                Dpdt_Value = 0
                                number_of_Results = 1
                                Reset_Time = 0
                                Expected_Result[0] = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Checksum_value'] 
                                Exec_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Checksum_time']
                                Expected_DTC = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Checksum_DTC']
                                Exec_Type = "TYPE_A"
                                Canid_Master_CANape_FailSafe_Path = Master_CANape_FailSafe_Path + "\\ID" + CANID_Failsafe
                                dest_FailSafe_Result_Folder = FailSafe_Result_Folder[Failsafe_Variant_Increment] + "\\ID" + CANID_Failsafe + "\\" + FailSafe_Category_Pre
                                copy_folder(Canid_Master_CANape_FailSafe_Path,dest_FailSafe_Result_Folder)

                                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                                    format='%(asctime)s - %(levelname)s - %(message)s')             
                                logging.info('Failsafe Test Procedure For Check sum Started')                                                                 # Logging info in the log file                                
                                count_hyperlinks,count_hyperlinks_CA = Fail_safe_Test_procedure(Inp_Sys_Var,Sys_Var_Set,Dependent_signal,Dpdt_Signal_Set,Exec_Type, \
                                                         Dest_Procedure_failsafe_result,FailSafe_Category_Pre,Dpdt_Value,Reset_Value,Input_Value, \
                                                         Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                                         Expected_Result,count_hyperlinks,Expected_DTC,FailSafe_Result_Folder[Failsafe_Variant_Increment],word,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                                         check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list, \
                                                         Open_Test_case_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,Multiple_Input_signal_After_list)
                                #make one while loop which monitor delay_variable to 4

                            elif Failsafe_categary_list[failsafe_Categate_loop] == "JT1" and CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_value']  != '':
                                Volt_Ref = 0
                                Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(CANONOFF_Inp_Sys_Var_row,6).value
                                print Inp_Sys_Var
                                Sys_Var_Set = None
                                Dpdt_Signal_Set = None
                                Dependent_signal = None
                                Input_Value = 0
                                Reset_Value = 1
                                Dpdt_Value = 0
                                number_of_Results = 2
                                Expected_Result[0] = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_value']
                                Expected_Result[1] = '0'
                                set_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_time'] 
                                Exec_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_time']
                                Reset_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT2_time']
                                Expected_DTC = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_DTC']
                                Exec_Type = "TYPE_B"
                                Canid_Master_CANape_FailSafe_Path = Master_CANape_FailSafe_Path + "\\ID" + CANID_Failsafe
                                dest_FailSafe_Result_Folder = FailSafe_Result_Folder[Failsafe_Variant_Increment] + "\\ID" + CANID_Failsafe + "\\" + FailSafe_Category_Pre
                                copy_folder(Canid_Master_CANape_FailSafe_Path,dest_FailSafe_Result_Folder)

                                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                                    format='%(asctime)s - %(levelname)s - %(message)s')             
                                logging.info('Failsafe Test Procedure For JT1 Started')                                                                 # Logging info in the log file                                
                                count_hyperlinks,count_hyperlinks_CA = Fail_safe_Test_procedure(Inp_Sys_Var,Sys_Var_Set,Dependent_signal,Dpdt_Signal_Set,Exec_Type, \
                                                         Dest_Procedure_failsafe_result,FailSafe_Category_Pre,Dpdt_Value,Reset_Value,Input_Value, \
                                                         Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                                         Expected_Result,count_hyperlinks,Expected_DTC,FailSafe_Result_Folder[Failsafe_Variant_Increment],word,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                                         check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list, \
                                                         Open_Test_case_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,Multiple_Input_signal_After_list)
                                #make one while loop which monitor delay_variable to 4
                            
                            elif Failsafe_categary_list[failsafe_Categate_loop] == "JT2" and CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT2_value']  != '':
                                Volt_Ref = 0
                                Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(CANONOFF_Inp_Sys_Var_row,6).value
                                print Inp_Sys_Var
                                Sys_Var_Set = None                            
                                Dpdt_Signal_Set = None
                                Dependent_signal = None
                                Input_Value = 0
                                Reset_Value = 1
                                Dpdt_Value = 0
                                number_of_Results = 3                            
                                Expected_Result[0] = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT2_value']
                                Expected_Result[1] = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_value']
                                Expected_Result[2] = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT2_value']
                                Expected_DTC = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT2_DTC']
                                set_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT1_time'] 
                                Exec_Time = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['JT2_time']
                                Reset_Time = 2000
                                Exec_Type = "TYPE_C"
                                Canid_Master_CANape_FailSafe_Path = Master_CANape_FailSafe_Path + "\\ID" + CANID_Failsafe
                                dest_FailSafe_Result_Folder = FailSafe_Result_Folder[Failsafe_Variant_Increment] + "\\ID" + CANID_Failsafe + "\\" + FailSafe_Category_Pre
                                copy_folder(Canid_Master_CANape_FailSafe_Path,dest_FailSafe_Result_Folder)
                                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                                    format='%(asctime)s - %(levelname)s - %(message)s')             
                                logging.info('Failsafe Test Procedure For JT2 Started')                                                                 # Logging info in the log file                                
                                count_hyperlinks,count_hyperlinks_CA = Fail_safe_Test_procedure(Inp_Sys_Var,Sys_Var_Set,Dependent_signal,Dpdt_Signal_Set,Exec_Type, \
                                                         Dest_Procedure_failsafe_result,FailSafe_Category_Pre,Dpdt_Value,Reset_Value,Input_Value, \
                                                         Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                                         Expected_Result,count_hyperlinks,Expected_DTC,FailSafe_Result_Folder[Failsafe_Variant_Increment],word,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                                         check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list, \
                                                        Open_Test_case_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,Multiple_Input_signal_After_list)
                                #make one while loop which monitor delay_variable to 4
                                
                            elif Failsafe_categary_list[failsafe_Categate_loop] == "Failsafe_Signals" and CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Failsafe_Signals']  != '':
                                Volt_Ref = 0
                                var_dept_signal_inc = 1
                                Dpdt_Signal_Set = None
                                Dependent_signal = None
                                Result_Dependent_Signal_Col = 20
                                Result_Dependent_Signal_Value_Col = 22
                                Result_Dependent_Signal_Comparison_Col = 21
                                Input_Value = 0
                                Reset_Value = 1
                                Dpdt_Value = 0
                                number_of_Results = 3  
                                Row_Sig = 0
                                Test_Case_col_number = 5
                                row_tp = 0
                                Reset_Time = 2000
                                Multiple_Input_signal_col = 23
                                words  = CanID_Data_FailSafe_Dict[VehicleName][Variant_Name][Ecu_Name_curItem][CANID_Failsafe]['Failsafe_Signals'].split(",")   
                                print words,"words"
                                print "Entered into failsafe if loop"
                                for word in words:
                                    print "word",word
                                    FAILSAFE_CAT_entry["state"] = NORMAL #bhau check it
                                    FAILSAFE_CAT_entry.delete(0,END)
                                    FAILSAFE_CAT_entry.insert(0,word)
                                    FAILSAFE_CAT_entry["state"] = DISABLED
                                    for row_tp in range (1,DIMPSheet_Failsafe_Failsafe_List_Row):
                                        print "row_tp",row_tp
                                        print "Test_Case_col_number",Test_Case_col_number
                                        print DIMPSheet_Failsafe_Failsafe_List.cell(row_tp,Test_Case_col_number).value
                                        if DIMPSheet_Failsafe_Failsafe_List.cell(row_tp,Test_Case_col_number).value == "ID" + CANID_Failsafe + '_' + word:
                                            Row_Sig = row_tp
                                            print "hi",Row_Sig
                                            break;
                                    same_signal_check = False
                                    Another_Signal_Name = 'Not_Possible'
                                    if 'AnotherSignals' in word: #gan did change
                                        Another_Signal_Name = word.rsplit('_',1)[1]
                                        word = word.rsplit('_',1)[0]
                                        same_signal_check = True                                        
                                    #*********************To find the row number and path of the failsafe signal input system variable*****************************#
                                    for row_Failsafe_signal_test_row in range (0,FailSafe_Test_Procedure_sheet.nrows):
                                        if FailSafe_Test_Procedure_sheet.cell(row_Failsafe_signal_test_row,4).value == "ID" + CANID_Failsafe + "_" + word + "_" + "Set":
                                            print "ID" + CANID_Failsafe + "_" + word + "_" + "Set"
                                            print "Enterted into master failsafe signal check set"
                                            Failsafe_Signal_Sys_Var_Set_row = row_Failsafe_signal_test_row
                                        elif FailSafe_Test_Procedure_sheet.cell(row_Failsafe_signal_test_row,4).value == "ID" + CANID_Failsafe + "_" + word + "_" + "Value":
                                            Failsafe_Signal_Sys_Var_Value_row = row_Failsafe_signal_test_row
                                            print "Enterted into master failsafe signal check set"
                                    print "Failsafe_Signal_Sys_Var_Set_row",Failsafe_Signal_Sys_Var_Set_row
                                    print "Failsafe_Signal_Sys_Var_Value_row",Failsafe_Signal_Sys_Var_Value_row
                                    Inp_Sys_Var = FailSafe_Test_Procedure_sheet.cell(Failsafe_Signal_Sys_Var_Value_row,6).value #value path
                                    Sys_Var_Set = FailSafe_Test_Procedure_sheet.cell(Failsafe_Signal_Sys_Var_Set_row,6).value   #Set path
                                    print "Input_sys_var_Set_path",Sys_Var_Set
                                    print "Input_sys_var_Value_path",Inp_Sys_Var
                                    #*********************************************************************************************************************************#
                                    Result_Dependent_Signal_list = []
                                    Result_Dependent_Signal_list_value = []
                                    Result_Dependent_Signal_Comparison_list_value = []
                                    Multiple_Input_signal_list =[]
                                    var_dept_signal_inc = 1
                                    Input_Sig_Col = 5;
                                    Dependent_Signal_Col = 7;
                                    Dependent_Signal_Set_Col = 8; 
                                    Dependent_Signal_Value_Col = 9;
                                    Set_Time_Col = 10;
                                    Reset_Time_Col = 11;
                                    Signal_set_Value_Col = 12;
                                    Signal_reset_Value_Col = 13;
                                    exec_Time_Col = 14;
                                    Expt_Result_Col = 16;
                                    Expt_Result2_Col = 17;
                                    Type_Col = 18;
                                    Expected_dtc_col = 19;
                                    Exec_Time=0 #gan
                                    Multiple_Input_signal_col = 23
                                    Multiple_Input_signal_After_col = 24
                                    Check_Multiple_Input_Signal = False
                                    Check_Multiple_Input_Signal_After = False
                                    
                                    print "Row_Sig",Row_Sig
                                    Input_Value = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Signal_set_Value_Col).value;
                                    print "Input_Value",Input_Value
                                    Reset_Value = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Signal_reset_Value_Col).value;
                                    print "Reset_Value",Reset_Value
                                    Dpdt_Signal_Set = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Dependent_Signal_Col).value;
                                    print "Dpdt_Signal_Set",Dpdt_Signal_Set 
                                    Dpdt_Value = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Dependent_Signal_Value_Col).value;
                                    print "Dpdt_Value",Dpdt_Value
                                    Expected_Result[0] = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Expt_Result_Col).value;
                                    print "Expected_Result[0]",Expected_Result[0]
                                    Expected_Result[1] = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Expt_Result2_Col).value;
                                    print "Expected_Result[1]",Expected_Result[1]
                                    Exec_Time = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, exec_Time_Col).value;
                                    print "Exec_Time",Exec_Time
                                    set_Time = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Set_Time_Col).value;
                                    print "set_Time",set_Time
                                    Reset_Time = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Reset_Time_Col).value;
                                    print "Reset_Time",Reset_Time
                                    Exec_Time_2 = float(Exec_Time) + 3;
                                    print "Exec_Time_2",Exec_Time_2
                                    Exec_Type = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Type_Col).value;
                                    print "Exec_Type",Exec_Type
                                    Expected_DTC = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Expected_dtc_col).value;
                                    testCaseName = Exec_Type;
                                    Canid_Master_CANape_FailSafe_Path = Master_CANape_FailSafe_Path + "\\ID" + CANID_Failsafe
                                    print "gan",word
                                    #*********************************************Extraction of data for TypeD procedure starts here******************************************************************************#
                                    if DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Multiple_Input_signal_col).value != '':
                                        Check_Multiple_Input_Signal = True
                                        Multiple_Input_signal_list = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Multiple_Input_signal_col).value.split(',')
                                        print "Multiple_Input_signal_list ",Multiple_Input_signal_list
                                    if DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Multiple_Input_signal_After_col).value != '':
                                        Check_Multiple_Input_Signal_After = True
                                        Multiple_Input_signal_After_list = DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Multiple_Input_signal_After_col).value.split(',')
                                        print "Multiple_Input_signal_After_list ",Multiple_Input_signal_After_list
                                        
                                    #*********************************************Extraction of data for TypeD procedure starts here******************************************************************************#                                    
                                    #*********************************************Multiple depedent signal judgement ganpi*****************************************************************************#
                                    if DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Result_Dependent_Signal_Col).value != '':
                                        check_dependent_signal = True 
                                        Result_Dependent_Signal_list.append(DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Result_Dependent_Signal_Col).value)                                       #Extract first depedent signal 
                                        Result_Dependent_Signal_Comparison_list_value.append(DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig, Result_Dependent_Signal_Comparison_Col).value)                                       #Extract first depedent signal 
                                        Result_Dependent_Signal_list_value.append(DIMPSheet_Failsafe_Failsafe_List.cell(Row_Sig,Result_Dependent_Signal_Value_Col).value)                           #Extract first depedent signal value
                                        print DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig + var_dept_signal_inc), Input_Sig_Col).value
                                        while (DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig + var_dept_signal_inc), Input_Sig_Col).value  == '' ):                                                 #check if any other dependent signal is there,If so extract those signals
                                            print "while loop"
                                            if '-' in DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig + var_dept_signal_inc), Result_Dependent_Signal_Col).value:
                                                temp_judgment_signal = DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig + var_dept_signal_inc), Result_Dependent_Signal_Col).value.split('-')
                                                Result_Dependent_Signal_list.extend([temp_judgment_signal[0],temp_judgment_signal[1]])
                                                Result_Dependent_Signal_list_value.extend([DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig+ var_dept_signal_inc), Result_Dependent_Signal_Value_Col ).value, \
                                                                                          DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig+ var_dept_signal_inc), Result_Dependent_Signal_Value_Col ).value])
                                                Result_Dependent_Signal_Comparison_list_value.extend(['-', \
                                                                                                     DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig+ var_dept_signal_inc), Result_Dependent_Signal_Comparison_Col ).value])                                                
                                            else:
                                                Result_Dependent_Signal_list.append(DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig + var_dept_signal_inc), Result_Dependent_Signal_Col).value)          
                                                Result_Dependent_Signal_list_value.append(DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig+ var_dept_signal_inc), Result_Dependent_Signal_Value_Col ).value)
                                                Result_Dependent_Signal_Comparison_list_value.append(DIMPSheet_Failsafe_Failsafe_List.cell((Row_Sig+ var_dept_signal_inc), Result_Dependent_Signal_Comparison_Col ).value)
                                            var_dept_signal_inc = var_dept_signal_inc + 1
                                    #****************************************************************************************************************************************************************#
                                    print "Result_Dependent_Signal_list",Result_Dependent_Signal_list
                                    print "Result_Dependent_Signal_list_value",Result_Dependent_Signal_list_value
                                    print "Result_Dependent_Signal_Comparison_list_value",Result_Dependent_Signal_Comparison_list_value
                                    if same_signal_check == True:                                    
                                        dest_FailSafe_Result_Folder = FailSafe_Result_Folder[Failsafe_Variant_Increment] + "\\ID" + CANID_Failsafe + "\\" + FailSafe_Category_Pre + "\\" + word + '_' + Another_Signal_Name
                                    else:
                                        dest_FailSafe_Result_Folder = FailSafe_Result_Folder[Failsafe_Variant_Increment] + "\\ID" + CANID_Failsafe + "\\" + FailSafe_Category_Pre + "\\" + word                                        
                                    print Canid_Master_CANape_FailSafe_Path
                                    copy_folder(Canid_Master_CANape_FailSafe_Path,dest_FailSafe_Result_Folder)
                                    count_hyperlinks,count_hyperlinks_CA = Fail_safe_Test_procedure(Inp_Sys_Var,Sys_Var_Set,Dependent_signal,Dpdt_Signal_Set,Exec_Type, \
                                                             Dest_Procedure_failsafe_result,FailSafe_Category_Pre,Dpdt_Value,Reset_Value,Input_Value, \
                                                             Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                                             Expected_Result,count_hyperlinks,Expected_DTC,FailSafe_Result_Folder[Failsafe_Variant_Increment],word,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                                             check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list, \
                                                             Open_Test_case_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,Multiple_Input_signal_After_list)                        
                            
                            FAILSAFE_overall_progressbar["value"] = overall_progressbar_value
                            overall_progressbar_value += 1
##                        overall_progressbar_value += 1
##                        FAILSAFE_overall_progressbar["value"] = overall_progressbar_value
                Failsafe_Variant_Increment = Failsafe_Variant_Increment + 1
                FAILSAFE_progressbar["value"] = 14                        
        except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.exception('Test case execution stopped abrubtly')
                    
                

                
    def Fail_safe_Test_procedure(Inp_Sys_Var,Sys_Var_Set,Dependent_signal,Dpdt_Signal_Set,Exec_Type,Dest_Procedure_failsafe_result, \
                                 FailSafe_Category_Pre,Dpdt_Value,Reset_Value,Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder, \
                                 CANID_Failsafe,Message_counter_Inp_Sys_Var,Expected_Result,count_hyperlinks,Expected_DTC,Result_FailSafe_Result_Folder,word,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                 check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                 Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,Multiple_Input_signal_After_list):
        global myAppl,varaint_path
        print "Hi, I am in failsafe Test procedure function"
        print "To edit"
        print "Exec_Type"
        print Exec_Type
        SignalData = []
        SigInfo = []
        sig_data_sheet_Failsafe_sheet = ' '
        SignalData_After = []
        SigInfo_After = []
        sig_data_sheet_Failsafe_sheet_After = ' '
        print "FailSafe_Category_Pre"
        print "Sys_Var_Set",Sys_Var_Set
        print "Inp_Sys_Var",Inp_Sys_Var
        print FailSafe_Category_Pre
        CANapeInpSysVar = Inp_Sys_Var
        print "CANapeInpSysVar",CANapeInpSysVar
        FAILSAFE_progressbar["value"] = 6
        Result_FailSafe_Result_Folder = Result_FailSafe_Result_Folder + "\\" + "00_Screenshots"
        write_ss_info = open(dest_FailSafe_Result_Folder + "\\Screen_Shot.txt","w")
        if not (os.path.exists(Result_FailSafe_Result_Folder)):
            os.mkdir(Result_FailSafe_Result_Folder)
        write_ss_info.write(Result_FailSafe_Result_Folder + "\n")
        write_ss_info.write(CANID_Failsafe)
        write_ss_info.write("_")
        write_ss_info.write(FailSafe_Category_Pre)        
        if word != '':
            write_ss_info.write("_")
            write_ss_info.write(word)
            Screen_shot_path = Result_FailSafe_Result_Folder + "\\" + CANID_Failsafe + "_" + FailSafe_Category_Pre + "_" + word + '_1' + ".jpg"
        else:
            Screen_shot_path = Result_FailSafe_Result_Folder + "\\" + CANID_Failsafe + "_" + FailSafe_Category_Pre + '_1' + ".jpg"
        write_ss_info.close()
        
        if Sys_Var_Set != None:
            CANapeSysVarSet = Sys_Var_Set
        if Dependent_signal != None:
            Dependent_signal_List = Dependent_signal.split(";")
            print Dependent_signal_List[0]
        if Dpdt_Signal_Set != None:
            Dpdt_Signal_Set_List = Dpdt_Signal_Set.split(";")
            print Dpdt_Signal_Set_List[0]

        if Check_Multiple_Input_Signal == True:
            SignalData,SigInfo,sig_data_sheet_Failsafe_sheet = Failsafe_Test_Procedure_TypeD(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value,Reset_Value, \
                                  Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                  Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                  check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet)

        if Check_Multiple_Input_Signal_After == True:
            SignalData_After,SigInfo_After,sig_data_sheet_Failsafe_sheet_After = Failsafe_Test_Procedure_TypeD(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value,Reset_Value, \
                                  Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                  Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                  check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_After_list,Open_Test_case_sheet)

            
        if Exec_Type == "TYPE_A":
            print "Entered to failsafe type A"
                
            count_hyperlinks,count_hyperlinks_CA = Failsafe_Test_Procedure_TypeA(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value, \
                                          Reset_Value,Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe, \
                                          Message_counter_Inp_Sys_Var,Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder ,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                          check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                                                                 SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,SignalData_After,SigInfo_After, \
                                                                                 sig_data_sheet_Failsafe_sheet_After)
            time.sleep(1)
        elif Exec_Type == "TYPE_B":
            print "Entered to failsafe type B"
            count_hyperlinks,count_hyperlinks_CA = Failsafe_Test_Procedure_TypeB(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value, \
                                          Reset_Value,Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe, \
                                          Message_counter_Inp_Sys_Var,Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                          check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                                                                 SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,SignalData_After,SigInfo_After, \
                                                                                 sig_data_sheet_Failsafe_sheet_After)
            time.sleep(1)            
        elif Exec_Type == "TYPE_C":
            print "Entered to failsafe type C"              
            count_hyperlinks,count_hyperlinks_CA = Failsafe_Test_Procedure_TypeC(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value, \
                                          Reset_Value,Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe, \
                                          Message_counter_Inp_Sys_Var,Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                          check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                                                                 SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,SignalData_After,SigInfo_After, \
                                                                                 sig_data_sheet_Failsafe_sheet_After)
            time.sleep(1)
        print count_hyperlinks
        print count_hyperlinks_CA
        print "count_hyperlinks"        
        
##        MDF_Files_Path = Dest_Procedure_failsafe_result
##        MDF_sourceFile
##        MDF_DestFile
##        Graphics_Path
        return count_hyperlinks,count_hyperlinks_CA
        



    def Failsafe_Test_Procedure_TypeA(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value,Reset_Value, \
                                      Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                      Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                      check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                      SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,SignalData_After,SigInfo_After, \
                                      sig_data_sheet_Failsafe_sheet_After):
        global myAppl,Actual_DTC_subArray_set,Actual_DTC_set,varaint_path,DTC_string_path,DTC_string_path_1
        Actual_DTC_set = ['','','','','','','','','','']
        Actual_DTC_subArray_set =  ['','','','','','','','','','']
        FAILSAFE_progressbar["value"] = 7
        pathTextFile = dest_FailSafe_Result_Folder
        Dspace_Trace_path = dest_FailSafe_Result_Folder + "\\Dspace_Trace.txt" 
        pathTextFile = pathTextFile + "\\" + "Sync.txt"
        #Signal_data_sheet_name = "Signal_Data"
        #Signal_data_sheet = Open_Test_case_sheet.sheet_by_name(Signal_data_sheet_name)
        
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(varaint_path).Write(3)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(1)        
        if "VoltCheck" in FailSafe_Category_Pre:
            print "VoltCheck We need to check"
            #Start_CANape(dest_FailSafe_Result_Folder)    #Start Canape for Failsafe
            print "we need to start measurement also"
            time.sleep(1)
        else:
            print "canape start"
            Start_CANape(dest_FailSafe_Result_Folder)    #Start Canape for Failsafe
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(1)        
        flag_start = 0
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Capal Script Started For TYPA')                                                                 # Logging info in the log file

        FAILSAFE_progressbar["value"] = 8
        while(flag_start == 0):                    
            syncFileRead = open(pathTextFile,'r')
            valueRead = syncFileRead.read()
            syncFileRead.close()
            if (valueRead == '7'):
                time.sleep(1);
                flag_start = 1
        if Check_Multiple_Input_Signal == True:
            ADAS_HILS_FAILSAFE_AUTOMATION(SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,myAppl)
        if Dependent_signal != None:
            myAppl.Variable(CANapeInpSysVar).Write(Dpdt_Value)
            time.sleep(1)
            if Dpdt_Signal_Set == None:
                myAppl.Variable(Dpdt_Signal_Set).Write(1)
        time.sleep(3)
        myAppl.Variable(CANapeInpSysVar).Write(Input_Value)         # Write Input_value
        if Sys_Var_Set != None:
            myAppl.Variable(Sys_Var_Set).Write(1)
        time_sleep = 4 + float(Exec_Time)
        time.sleep(time_sleep)
##        DTC_String ='Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray1}'
##        DTC_String_1 ='Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray2}'
        DTC_string_path_split =DTC_string_path.rindex("{")
        DTC_string_path_temp = DTC_string_path[:DTC_string_path_split-1] 
        #******************dtc check********************************#
##        Actual_bit_set = []
##        Actual_bit_set_status = []
##        path_variable_read_DTC = 'Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray1}'       
##        for DTC_bit_value in range(1,11):
##            path_variable_read_DTC = 'Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC' + str(DTC_bit_value) + '{SubArray1}'
##            print path_variable_read_DTC
##            status_Read_DTC = 'Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC' + str(DTC_bit_value) + '{SubArray2}'
##            print status_Read_DTC
##            Actual_bit_set.append(myAppl.Variable(path_variable_read_DTC).Read())
##            print myAppl.Variable(status_Read_DTC).Read()
##            Actual_bit_set_status.append(myAppl.Variable(status_Read_DTC).Read())
##            print myAppl.Variable(status_Read_DTC).Read()
####        Actual_bit_set = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray1}').Read()
####        Actual_bit_set_current=myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC4{SubArray1}').Read()
####        Actual_bit_set1 = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray2}').Read()
##        myAppl.Variable(varaint_path).Write(3)
##        #Actual_bit_set = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray2}').Read()
##        print "Actual_bit_set"
##        print Actual_bit_set
##        print Actual_bit_set_status
##        #****************************ends here***************************#
        if Sys_Var_Set != None:
            myAppl.Variable(Sys_Var_Set).Write(0)
        myAppl.Variable(CANapeInpSysVar).Write(Reset_Value)
        if Check_Multiple_Input_Signal_After == True:
            ADAS_HILS_FAILSAFE_AUTOMATION(SignalData_After,SigInfo_After,sig_data_sheet_Failsafe_sheet_After,myAppl)        
        time.sleep(3)
        myAppl.Variable(varaint_path).Write(2)
        time.sleep(.5)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(.5)
        for dtc_array_count in range(0,9):
            Actual_DTC_set[dtc_array_count] = myAppl.Variable(str(DTC_string_path_temp + str(dtc_array_count + 1) + "{SubArray1}")).Read()
            Actual_DTC_subArray_set[dtc_array_count] = myAppl.Variable(str(DTC_string_path_temp + str(dtc_array_count + 1) + "{SubArray2}")).Read()
            
##        Actual_DTC_set[0] = myAppl.Variable(DTC_String).Read()
##        print "Actual_DTC_set[0]",Actual_DTC_set[0]
##        Actual_DTC_set[1] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC2{SubArray1}').Read()
##        Actual_DTC_set[2] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC3{SubArray1}').Read()
##        Actual_DTC_set[3] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC4{SubArray1}').Read()
##        Actual_DTC_set[4] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC5{SubArray1}').Read()
##        Actual_DTC_set[5] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC6{SubArray1}').Read()
##        Actual_DTC_set[6] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC7{SubArray1}').Read()
##        Actual_DTC_set[7] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC8{SubArray1}').Read()
##        Actual_DTC_set[8] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC9{SubArray1}').Read()
##        Actual_DTC_set[9] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC10{SubArray1}').Read()
##        Actual_DTC_subArray_set[0] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray2}').Read()
##        Actual_DTC_subArray_set[1] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC2{SubArray2}').Read()
##        Actual_DTC_subArray_set[2] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC3{SubArray2}').Read()
##        Actual_DTC_subArray_set[3] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC4{SubArray2}').Read()
##        Actual_DTC_subArray_set[4] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC5{SubArray2}').Read()
##        Actual_DTC_subArray_set[5] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC6{SubArray2}').Read()
##        Actual_DTC_subArray_set[6] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC7{SubArray2}').Read()
##        Actual_DTC_subArray_set[7] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC8{SubArray2}').Read()
##        Actual_DTC_subArray_set[8] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC9{SubArray2}').Read()
##        Actual_DTC_subArray_set[9] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC10{SubArray2}').Read()
        print Actual_DTC_set
        print Actual_DTC_subArray_set
##        myAppl.Variable(varaint_path).Write(3)        
        syncFileWrite = open(pathTextFile,'w')
        sync_num = 9
        valueWrite= str(sync_num)
        syncFileWrite.write(valueWrite)
        syncFileWrite.close()        
        if Check_Multiple_Input_Signal == True:
        
            for m in range(1,sig_data_sheet_Failsafe_sheet.nrows):

                set_sig_reset_value = None
                set_sig_path = None
                set_sig_default_value = None
                set_sig_Appl = None
                set_sig_reset_value = int(sig_data_sheet_Failsafe_sheet.cell(m,3).value)
                set_sig_path = sig_data_sheet_Failsafe_sheet.cell(m,1).value
                set_sig_default_value = sig_data_sheet_Failsafe_sheet.cell(m,2).value
                set_sig_Appl = sig_data_sheet_Failsafe_sheet.cell(m,4).value
                if set_sig_reset_value ==1 :
                    if (set_sig_Appl == 'All'):
                        try:
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(1)
                        except:
                            pass
        if Check_Multiple_Input_Signal_After == True:
        
            for m in range(1,sig_data_sheet_Failsafe_sheet_After.nrows):

                set_sig_reset_value = None
                set_sig_path = None
                set_sig_default_value = None
                set_sig_Appl = None
                set_sig_reset_value = int(sig_data_sheet_Failsafe_sheet_After.cell(m,3).value)
                set_sig_path = sig_data_sheet_Failsafe_sheet_After.cell(m,1).value
                set_sig_default_value = sig_data_sheet_Failsafe_sheet_After.cell(m,2).value
                set_sig_Appl = sig_data_sheet_Failsafe_sheet_After.cell(m,4).value
                if set_sig_reset_value ==1 :
                    if (set_sig_Appl == 'All'):
                        try:
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(1)
                        except:
                            pass

        flag_start = 0
        while(flag_start == 0):
                    
            syncFileRead = open(pathTextFile,'r')

            valueRead = syncFileRead.read()
            syncFileRead.close()

            if (valueRead == '8'):
                time.sleep(1);
                flag_start = 1
##        #*****************Ramscope popup clearance*****************#
##        hwndMain = 0
##        tend = time.time() + 5
##        while time.time() < tend:        
##            hwndMain = win32gui.FindWindow(None,"Vector CANape")
##            print hwndMain
######        if hwndMain != 0 :
##            try:           
##                win32gui.SetForegroundWindow(hwndMain)   
##                win32api.PostMessage(hwndMain,win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)                
##            except:
##                print "hwndMain",hwndMain
##                win32api.PostMessage(hwndMain,win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
##        #*****************Ramscope popup clearance*****************#
        myAppl.Variable(Power_Supply_path).Write(129)
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(varaint_path).Write(3)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Procedure for measurement completed')                                                                 # Logging info in the log file

        os.remove(dest_FailSafe_Result_Folder + "\\CANape_Script_V4.scr")
        FAILSAFE_progressbar["value"] = 9
        count_hyperlinks,count_hyperlinks_CA = Judgement_final(dest_FailSafe_Result_Folder + "\\CANape.txt",CANapeInpSysVar,CANID_Failsafe,Exec_Time,temp_Master_Result_Report_FailSafe,FailSafe_Category_Pre, \
                        Message_counter_Inp_Sys_Var,Expected_Result,Exec_Time,count_hyperlinks,Expected_DTC,Exec_Type,Input_Value,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var, \
                        count_hyperlinks_CA,dest_FailSafe_Result_Folder,check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Dspace_Trace_path)
        time.sleep(3)
        print count_hyperlinks
        print "count_hyperlinks"        
        FAILSAFE_progressbar["value"] = 14
        return count_hyperlinks,count_hyperlinks_CA

  
    def Failsafe_Test_Procedure_TypeB(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value,Reset_Value, \
                                      Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                      Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA,check_dependent_signal,Result_Dependent_Signal_list, \
                                      Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                      SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,SignalData_After,SigInfo_After, \
                                      sig_data_sheet_Failsafe_sheet_After):
        global myAppl,Actual_DTC_set,Actual_DTC_subArray_set,varaint_path,DTC_string_path,DTC_string_path_1
        FAILSAFE_progressbar["value"] = 7
        Actual_DTC_set = ['','','','','','','','','','']
        Actual_DTC_subArray_set =  ['','','','','','','','','','']        
        Failsafe_Threshold_Time = 10
        pathTextFile = dest_FailSafe_Result_Folder
        Dspace_Trace_path = dest_FailSafe_Result_Folder + "\\Dspace_Trace.txt" 
        pathTextFile = pathTextFile + "\\" + "Sync.txt"
    
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(.5)
        myAppl.Variable(varaint_path).Write(3)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(.5)
        Start_CANape(dest_FailSafe_Result_Folder)    #Start Canape for Failsafe        

        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(.5)
        myAppl.Variable(Power_Supply_path).Write(1)        
        flag_start = 0
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Capal Script Started For TYPB')                                                                 # Logging info in the log file

        FAILSAFE_progressbar["value"] = 8
        while(flag_start == 0):                    
            syncFileRead = open(pathTextFile,'r')
            valueRead = syncFileRead.read()
            syncFileRead.close()
            if (valueRead == '7'):
                time.sleep(1);
                flag_start = 1                    
            
        time.sleep(1)

        if Check_Multiple_Input_Signal == True:
            ADAS_HILS_FAILSAFE_AUTOMATION(SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,myAppl)

        if Dependent_signal != None:
            myAppl.Variable(CANapeInpSysVar).Write(Dpdt_Value)
            time.sleep(1)
            if Dpdt_Signal_Set == None:
                myAppl.Variable(Dpdt_Signal_Set).Write(1)

        #Failsafe Execution
        myAppl.Variable(Power_Supply_path).Write(1)            
        time.sleep(5)
        time_sleep = float(Exec_Time) * 2
        print time_sleep
        print Sys_Var_Set
        print Reset_Value
        print CANapeInpSysVar
        myAppl.Variable(CANapeInpSysVar).Write(Input_Value)         # Write Variant change value to variable        
        if Sys_Var_Set != None:
            myAppl.Variable(Sys_Var_Set).Write(1)            
            time.sleep(time_sleep)             
            myAppl.Variable(Sys_Var_Set).Write(0)
        else:
            time.sleep(time_sleep)
        myAppl.Variable(CANapeInpSysVar).Write(Reset_Value)
        if Check_Multiple_Input_Signal_After == True:
            ADAS_HILS_FAILSAFE_AUTOMATION(SignalData_After,SigInfo_After,sig_data_sheet_Failsafe_sheet_After,myAppl)                
##        myAppl.Variable(varaint_path).Write(2)
##        myAppl.Variable(varaint_path).Write(0)
        syncFileWrite = open(pathTextFile,'w')        
        sync_num = 9
        valueWrite= str(sync_num)
        syncFileWrite.write(valueWrite)
        syncFileWrite.close()
        if Check_Multiple_Input_Signal == True:
        
            for m in range(1,sig_data_sheet_Failsafe_sheet.nrows):

                set_sig_reset_value = None
                set_sig_path = None
                set_sig_default_value = None
                set_sig_Appl = None
                set_sig_reset_value = int(sig_data_sheet_Failsafe_sheet.cell(m,3).value)
                set_sig_path = sig_data_sheet_Failsafe_sheet.cell(m,1).value
                set_sig_default_value = sig_data_sheet_Failsafe_sheet.cell(m,2).value
                set_sig_Appl = sig_data_sheet_Failsafe_sheet.cell(m,4).value
                if set_sig_reset_value ==1 :
                    if (set_sig_Appl == 'All'):
                        try:
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(1)
                        except:
                            pass
        if Check_Multiple_Input_Signal_After == True:
        
            for m in range(1,sig_data_sheet_Failsafe_sheet_After.nrows):

                set_sig_reset_value = None
                set_sig_path = None
                set_sig_default_value = None
                set_sig_Appl = None
                set_sig_reset_value = int(sig_data_sheet_Failsafe_sheet_After.cell(m,3).value)
                set_sig_path = sig_data_sheet_Failsafe_sheet_After.cell(m,1).value
                set_sig_default_value = sig_data_sheet_Failsafe_sheet_After.cell(m,2).value
                set_sig_Appl = sig_data_sheet_Failsafe_sheet_After.cell(m,4).value
                if set_sig_reset_value ==1 :
                    if (set_sig_Appl == 'All'):
                        try:
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(1)
                        except:
                            pass

        
        flag_start = 0
        while(flag_start == 0):                        
            syncFileRead = open(pathTextFile,'r')
            valueRead = syncFileRead.read()
            syncFileRead.close()
            if (valueRead == '8'):
                time.sleep(1);
                flag_start = 1                        
        myAppl.Variable(Power_Supply_path).Write(129)
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(varaint_path).Write(3)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(5)
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Procedure for measurement completed')                                                                 # Logging info in the log file
        os.remove(dest_FailSafe_Result_Folder + "\\CANape_Script_V4.scr")
        FAILSAFE_progressbar["value"] = 9
        count_hyperlinks,count_hyperlinks_CA = Judgement_final(dest_FailSafe_Result_Folder + "\\CANape.txt",CANapeInpSysVar,CANID_Failsafe,Exec_Time,temp_Master_Result_Report_FailSafe,FailSafe_Category_Pre, \
                        Message_counter_Inp_Sys_Var,Expected_Result,Exec_Time,count_hyperlinks,Expected_DTC,Exec_Type,Input_Value,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var, \
                        count_hyperlinks_CA,dest_FailSafe_Result_Folder,check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Dspace_Trace_path)
        time.sleep(3)
        print count_hyperlinks
        print "count_hyperlinks"        
        FAILSAFE_progressbar["value"] = 14
        return count_hyperlinks,count_hyperlinks_CA
        
    def Failsafe_Test_Procedure_TypeC(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value,Reset_Value, \
                                      Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                      Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA,check_dependent_signal, \
                                      Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list,Open_Test_case_sheet, \
                                      SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,Check_Multiple_Input_Signal,Check_Multiple_Input_Signal_After,SignalData_After,SigInfo_After, \
                                      sig_data_sheet_Failsafe_sheet_After):
        global myAppl,Actual_DTC_set,Actual_DTC_subArray_set,varaint_path,DTC_string_path,DTC_string_path_1
        Actual_DTC_set = ['','','','','','','','','','']
        Actual_DTC_subArray_set =  ['','','','','','','','','','']        
        FAILSAFE_progressbar["value"] = 7
        pathTextFile = dest_FailSafe_Result_Folder
        Dspace_Trace_path = dest_FailSafe_Result_Folder + "\\Dspace_Trace.txt"         
        pathTextFile = pathTextFile + "\\" + "Sync.txt"
        print 'I riched here'
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(varaint_path).Write(3)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(1)        
        Start_CANape(dest_FailSafe_Result_Folder)    #Start Canape for Failsafe
        time.sleep(5)
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(1)        
        flag_start = 0
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Capal Script Started For TYPC')                                                                 # Logging info in the log file
        FAILSAFE_progressbar["value"] = 8
        while(flag_start == 0):                    
            syncFileRead = open(pathTextFile,'r')
            valueRead = syncFileRead.read()
            syncFileRead.close()
            if (valueRead == '7'):
                time.sleep(1);
                flag_start = 1                    
            
        time.sleep(1)
        if Check_Multiple_Input_Signal == True:
            ADAS_HILS_FAILSAFE_AUTOMATION(SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,myAppl)
        
        if Dependent_signal != None:
            myAppl.Variable(CANapeInpSysVar).Write(Dpdt_Value)
            time.sleep(1)
            if Dpdt_Signal_Set == None:
                myAppl.Variable(Dpdt_Signal_Set).Write(1)
        
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        time_sleep = 2 + float(Exec_Time)
        myAppl.Variable(CANapeInpSysVar).Write(Input_Value)         # Write Variant change value to variable        
        if Sys_Var_Set != None:
            myAppl.Variable(Sys_Var_Set).Write(1)
            time.sleep(time_sleep)
            myAppl.Variable(Sys_Var_Set).Write(0)
        else:
            time.sleep(time_sleep)
        myAppl.Variable(CANapeInpSysVar).Write(Reset_Value)
        if Check_Multiple_Input_Signal_After == True:
            ADAS_HILS_FAILSAFE_AUTOMATION(SignalData_After,SigInfo_After,sig_data_sheet_Failsafe_sheet_After,myAppl)                
        time.sleep(1)
        myAppl.Variable(varaint_path).Write(2)
        time.sleep(.5)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(.5)
        DTC_string_path_split =DTC_string_path.rindex("{")
        DTC_string_path_temp = DTC_string_path[:DTC_string_path_split-1]
        for dtc_array_count in range(0,9):
            Actual_DTC_set[dtc_array_count] = myAppl.Variable(str(DTC_string_path_temp + str(dtc_array_count + 1) + "{SubArray1}")).Read()
            Actual_DTC_subArray_set[dtc_array_count] = myAppl.Variable(str(DTC_string_path_temp + str(dtc_array_count + 1) + "{SubArray2}")).Read()
            
##        Actual_DTC_set[0] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray1}').Read()        
##        Actual_DTC_set[1] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC2{SubArray1}').Read()
##        Actual_DTC_set[2] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC3{SubArray1}').Read()
##        Actual_DTC_set[3] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC4{SubArray1}').Read()
##        Actual_DTC_set[4] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC5{SubArray1}').Read()
##        Actual_DTC_set[5] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC6{SubArray1}').Read()
##        Actual_DTC_set[6] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC7{SubArray1}').Read()
##        Actual_DTC_set[7] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC8{SubArray1}').Read()
##        Actual_DTC_set[8] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC9{SubArray1}').Read()
##        Actual_DTC_set[9] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC10{SubArray1}').Read()
##        Actual_DTC_subArray_set[0] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC1{SubArray2}').Read()
##        Actual_DTC_subArray_set[1] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC2{SubArray2}').Read()
##        Actual_DTC_subArray_set[2] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC3{SubArray2}').Read()
##        Actual_DTC_subArray_set[3] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC4{SubArray2}').Read()
##        Actual_DTC_subArray_set[4] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC5{SubArray2}').Read()
##        Actual_DTC_subArray_set[5] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC6{SubArray2}').Read()
##        Actual_DTC_subArray_set[6] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC7{SubArray2}').Read()
##        Actual_DTC_subArray_set[7] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC8{SubArray2}').Read()
##        Actual_DTC_subArray_set[8] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC9{SubArray2}').Read()
##        Actual_DTC_subArray_set[9] = myAppl.Variable('Model Root/Driver Block/CANdb set/DIAG/ID7C9_RX/PRC_READ_DTC/DTC10{SubArray2}').Read()
        print Actual_DTC_set
        print Actual_DTC_subArray_set
        syncFileWrite = open(pathTextFile,'w')
        sync_num = 9
        valueWrite= str(sync_num)
        syncFileWrite.write(valueWrite)        
        syncFileWrite.close()                                
        time.sleep(5)
        if Check_Multiple_Input_Signal == True or Check_Multiple_Input_Signal_After == True:
        
            for m in range(1,sig_data_sheet_Failsafe_sheet.nrows ):

                set_sig_reset_value = None
                set_sig_path = None
                set_sig_default_value = None
                set_sig_Appl = None
                set_sig_reset_value = int(sig_data_sheet_Failsafe_sheet.cell(m,3).value)
                set_sig_path = sig_data_sheet_Failsafe_sheet.cell(m,1).value
                set_sig_default_value = sig_data_sheet_Failsafe_sheet.cell(m,2).value
                set_sig_Appl = sig_data_sheet_Failsafe_sheet.cell(m,4).value
                if set_sig_reset_value ==1 :
                    if (set_sig_Appl == 'All'):
                        try:
                            myAppl.Variable(set_sig_path).Write(set_sig_default_value)
                            time.sleep(1)
                        except:
                            pass

        flag_start = 0
        while(flag_start == 0):
                        
            syncFileRead = open(pathTextFile,'r')
            valueRead = syncFileRead.read()
            syncFileRead.close()
            if (valueRead == '8'):
                time.sleep(1);
                flag_start = 1                        
        time.sleep(0.5)
        myAppl.Variable(Power_Supply_path).Write(129)
        myAppl.Variable(Power_Supply_path).Write(1)
        time.sleep(1)
        myAppl.Variable(varaint_path).Write(3)
        myAppl.Variable(varaint_path).Write(0)
        time.sleep(1)
        myAppl.Variable(Power_Supply_path).Write(129)
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Procedure for measurement completed')                                                                 # Logging info in the log file

        os.remove(dest_FailSafe_Result_Folder + "\\CANape_Script_V4.scr")
        FAILSAFE_progressbar["value"] = 9
        count_hyperlinks,count_hyperlinks_CA = Judgement_final(dest_FailSafe_Result_Folder + "\\CANape.txt",CANapeInpSysVar,CANID_Failsafe,Exec_Time,temp_Master_Result_Report_FailSafe, \
                        FailSafe_Category_Pre,Message_counter_Inp_Sys_Var,Expected_Result,Exec_Time,count_hyperlinks,Expected_DTC,Exec_Type,Input_Value,Result_FailSafe_Result_Folder, \
                        Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA,dest_FailSafe_Result_Folder,check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value, \
                                                               Result_Dependent_Signal_list_value,Dspace_Trace_path)
        time.sleep(3)
        print count_hyperlinks
        print "count_hyperlinks"
        FAILSAFE_progressbar["value"] = 14
        return count_hyperlinks,count_hyperlinks_CA
    
    def Failsafe_Test_Procedure_TypeD(Sys_Var_Set,CANapeInpSysVar,Dependent_signal,Dpdt_Signal_Set,FailSafe_Category_Pre,Dpdt_Value,Reset_Value, \
                                      Input_Value,Reset_Time,Exec_Time,dest_FailSafe_Result_Folder,CANID_Failsafe,Message_counter_Inp_Sys_Var, \
                                      Expected_Result,count_hyperlinks,Expected_DTC,Exec_Type,Result_FailSafe_Result_Folder,Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA, \
                                      check_dependent_signal,Result_Dependent_Signal_list,Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Multiple_Input_signal_list, \
                                      Open_Test_case_sheet):
        Signal_data_Failsafe_str = sig_data_sheet_str
        print "Enter into procedure TYPE_D"
        Failsafe_TypeD_Procedure_sheet = Open_Test_case_sheet.sheet_by_name(Multiple_Input_signal_list[0])
        sig_data_sheet_Failsafe_sheet = Open_Test_case_sheet.sheet_by_name(Signal_data_Failsafe_str)
        Failsafe_TypeD_Procedure_sheet_rows = Failsafe_TypeD_Procedure_sheet.nrows
        test_case_start_row = "test_case_no_" + Multiple_Input_signal_list[1]
        test_case_end_row = "end_test_case_" + Multiple_Input_signal_list[1]
        SignalData = []
        SigInfo = []
        SigNames = []
        execute_cont_str = ['exec_cont']
        exec_var_dep = ['exec_var_dep']
        exec_delay = ['exec_delay']
        exec_push_var_dep = ['exec_push_var_dep']
        exec_wait_var_dep = ['exec_wait_var_dep']
        execute_start_end_str = ['exec_start_end'] 
        sig_info={}
        
        for k in range(0, Failsafe_TypeD_Procedure_sheet_rows):
                    
            if Failsafe_TypeD_Procedure_sheet.cell(k,0).value== test_case_start_row :
                TestCase_Start_Row_failsafe= k
                break
                 
            else:
                continue
            
        for k in range(0, Failsafe_TypeD_Procedure_sheet_rows):
            if Failsafe_TypeD_Procedure_sheet.cell(k,0).value== test_case_end_row :
                TestCase_End_Row_failsafe = k
                break
            else:
                continue
        
        print TestCase_Start_Row_failsafe,TestCase_End_Row_failsafe,'Start End Row'

        save_var = 0
       
        for x in range(TestCase_Start_Row_failsafe + 1,TestCase_End_Row_failsafe):

            sig_name = Failsafe_TypeD_Procedure_sheet.cell(x,0).value
            sig_delay = Failsafe_TypeD_Procedure_sheet.cell(x,2).value
            for y in range(0,sig_data_sheet_Failsafe_sheet.nrows ):
                if sig_name == sig_data_sheet_Failsafe_sheet.cell(y, 0).value:
                    sig_path = sig_data_sheet_Failsafe_sheet.cell(y, 1).value
                    sig_value = sig_data_sheet_Failsafe_sheet.cell(y, 2).value
                    sig_reset = sig_data_sheet_Failsafe_sheet.cell(y, 3).value
            SigNames.append(sig_name)       
            sig_data = [sig_path,sig_value,sig_reset]                                                       # Collect all the signal data in list 'sig_data'
            sig_info[sig_name] = sig_data
            SigInfo.append(sig_info[sig_name])
             
           
            if str(Failsafe_TypeD_Procedure_sheet.cell(x,1).value) in \
               execute_start_end_str:                                                                       # If string in cloumn 2 is 'exec_start_end' then call 'get_data_execute_start_end_str' to collect test case data 
                signal_data = get_data_execute_start_end_str(sig_name,
                                                             sig_path,
                                                             Failsafe_TypeD_Procedure_sheet,
                                                             sig_delay,x,
                                                             sig_data_sheet_Failsafe_sheet)
                save_var = 1
            elif str(Failsafe_TypeD_Procedure_sheet.cell(x,1).value) in execute_cont_str:                                # If string in cloumn 2 is 'exec_cont' then call 'get_data_execute_conti' to collect test case data
                
                print "sig_name",sig_name
                print "sig_path",sig_path
                
                print "Failsafe_TypeD_Procedure_sheet",Failsafe_TypeD_Procedure_sheet
                print "sig_delay",sig_delay
                print "sig_data_sheet_Failsafe_sheet",sig_data_sheet_Failsafe_sheet
                print "x",x
              
                signal_data = get_data_execute_cont(sig_name,sig_path,
                                                    Failsafe_TypeD_Procedure_sheet,
                                                    sig_delay,x,
                                                    sig_data_sheet_Failsafe_sheet)
                save_var = 1    
            elif str(Failsafe_TypeD_Procedure_sheet.cell(x,1).value) in exec_var_dep:                                    # If string in cloumn 2 is 'exec_var_dep' then call 'get_data_execute_var_dep' to collect test case data   
                signal_data = get_data_execute_var_dep(sig_name,sig_path,
                                                       Failsafe_TypeD_Procedure_sheet,
                                                       sig_delay,x,
                                                       sig_data_sheet_Failsafe_sheet)
                save_var = 1
            elif str(Failsafe_TypeD_Procedure_sheet.cell(x,1).value) in exec_delay:                                      # If string in cloumn 2 is 'exec_delay' then call 'get_execute_delay' to collect test case data  
                signal_data = get_execute_delay(sig_name,sig_delay)
                save_var = 1


            SignalData.append(signal_data)
            
            if save_var == 1:                                                                               # Save_var = 1 signifies that atleast 1 signal is present in that test case 
                save_var = 0
                
            signal_data = []
            sig_info = {}
            
        print SignalData, 'SignalData'
        print SigInfo,'SigInfo'
        print sig_data_sheet_Failsafe_sheet,'sig_data_sheet_Failsafe_sheet'
        print myAppl,'myAppl'
        ##ADAS_HILS_FAILSAFE_AUTOMATION(SignalData,SigInfo,sig_data_sheet_Failsafe_sheet,myAppl)
        return (SignalData,SigInfo,sig_data_sheet_Failsafe_sheet)

    

    def Judgement_final(result_txt_file,Inp_Sys_Var,CANID_Failsafe,signal_time_period,temp_Master_Result_Report_FailSafe,FailSafe_Category_Pre, \
                        Message_counter_Inp_Sys_Var,Expected_Result,Exec_Time,count_hyperlinks,Expected_DTC,Exec_Type,Input_Value,Result_FailSafe_Result_Folder, \
                        Screen_shot_path,Checksum_Inp_Sys_Var,count_hyperlinks_CA,dest_FailSafe_Result_Folder,check_dependent_signal,Result_Dependent_Signal_list, \
                        Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,Dspace_Trace_path):
        global myAppl,Actual_DTC_subArray_set,Actual_DTC_set,same_signal_check,Output_Signal_Count,pathTextFile,First_screen_shot_time,Input_screen_shot_time,FailSafe_Display_Tree,Another_Signal_Name
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Judgment Started')                                                                 # Logging info in the log file
        test_result_sheet_canid_row_temp = -1
        Result_Dependent_Signal_Judgement = "CA"
        #DTC_Col_number = ['','','','','','','','']
        actual_bit_set = -1
        return_result_row_temp = -1
        afs_status_change_row_number = -1
        return_actual_bit_set = 0
        return_previos_time_stamp = 0
        time_Difference_On_Off = 0
        Actual_DTC_value = ""
        First_screen_shot_time = -150
        Input_screen_shot_time = -150
        afs_status_col_list = ['','','','','','','','','']
        afs_status_value_list = ['','','','','','','','','']
        afs_status_JT1_value_list = ['','','','','','','','','']
        afs_status_JT2_value_list = ['','','','','','','','','']
        dept_sig_num = 0
        
        Result_dependent_signal_Canape_col = range(len(Result_Dependent_Signal_list))
        excel= win32com.client.dynamic.Dispatch("Excel.Application")
        Report_FailSafe_write_WorkBook = excel.Workbooks.Open(str(temp_Master_Result_Report_FailSafe))
        Master_Result_Report_FailSafe_write_WorkBook = excel.Workbooks.Open(str(temp_Master_Result_Report_FailSafe))
        result_file_name_without_extension = result_txt_file.split('.')
        Number_of_xls_file = 0    
        result_xls_path = result_file_name_without_extension[0] + str(Number_of_xls_file) + '.xls'
        book = xlwt.Workbook()
        ws = book.add_sheet('canape')
        f = open(result_txt_file, 'r+')
        data = f.readlines()
        row_number_multi_xls = 1
        for i in range(len(data)):
            row = data[i].split(',')  # This will return a line of string data, you may need to convert to other formats depending on your use case        
            for j in range(len(row)):            
                ws.write(row_number_multi_xls, j, row[j])  # Write to cell i, j
            row_number_multi_xls = row_number_multi_xls + 1
            if (row_number_multi_xls  > 65000):
                print 'new excel created'
                book.save(str(result_xls_path))            
                Number_of_xls_file = Number_of_xls_file + 1
                result_xls_path = result_file_name_without_extension[0] + str(Number_of_xls_file) + '.xls'
                book = xlwt.Workbook()
                ws = book.add_sheet('canape')
                row_number_multi_xls = 1
        
        book.save(str(result_xls_path))
        f.close()
        Number_of_xls_file_result = 0
        result_xls_path = result_file_name_without_extension[0] + str(Number_of_xls_file_result) + '.xls'
        
        Result_FailSafe_WorkBook = xlrd.open_workbook(str(result_xls_path),formatting_info=True)
        Result_FailSafe_Canape_worksheet = Result_FailSafe_WorkBook.sheet_by_index(0)
        number_of_rows = Result_FailSafe_Canape_worksheet.nrows
        number_of_columns = Result_FailSafe_Canape_worksheet.ncols
        print "Good morning"
        if (FailSafe_Category_Pre == 'JT1' or  FailSafe_Category_Pre == 'JT2' ):
            if (Message_counter_Inp_Sys_Var != 'Path'):                
                temp_signal_name_inp_sys = Message_counter_Inp_Sys_Var.rsplit('/')[-3]
                signal_name_inp_sys = temp_signal_name_inp_sys[temp_signal_name_inp_sys.index('_')+len('_'):] 
            elif(Checksum_Inp_Sys_Var != 'Path'):
                temp_signal_name_inp_sys = Checksum_Inp_Sys_Var.rsplit('/')[-3]
                signal_name_inp_sys = temp_signal_name_inp_sys[temp_signal_name_inp_sys.index('_')+len('_'):] 
##            else:
##                return
        else:
            temp_signal_name_inp_sys = Inp_Sys_Var.rsplit('/')[-3]
            signal_name_inp_sys = temp_signal_name_inp_sys[temp_signal_name_inp_sys.index('_')+len('_'):] 
        
        if (FailSafe_Category_Pre == 'JT1' or  FailSafe_Category_Pre == 'JT2' ):
            signal_name_inp_sys = None
        print signal_name_inp_sys
        FAILSAFE_progressbar["value"] = 10 
        signal_input_sys_col = None
        for row_canape_result in range(number_of_rows):
            for col_canape_result in range(number_of_columns):
                ramscope_value  = Result_FailSafe_Canape_worksheet.cell(row_canape_result,col_canape_result).value            
                for afs_status_col_varaible in range(0,Output_Signal_Count):
                    if('.aFS_Status._' + str(afs_status_col_varaible) + '_[]' in ramscope_value):
                        afs_status_col_list[afs_status_col_varaible] = col_canape_result
                if signal_name_inp_sys  != None:
                    if(signal_name_inp_sys in ramscope_value and CANID_Failsafe in ramscope_value ):
                        signal_input_sys_col = col_canape_result
                #*****************Multiple signal judgement ganpi*********************************************#
                if(check_dependent_signal == True): #to find the respective column numbers of depedent signals in canape file
                    for dept_sig_num in range(len(Result_dependent_signal_Canape_col)):
                        if(Result_Dependent_Signal_list[dept_sig_num] in ramscope_value):
                            Result_dependent_signal_Canape_col[dept_sig_num] = col_canape_result                
                #***************************************************************************************#
        print Result_dependent_signal_Canape_col,"hi"
        print afs_status_col_list
        if FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2':
            excel= win32com.client.dynamic.Dispatch("Excel.Application")            
            JT1_JT2_workbook = excel.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
            JT1_JT2_worksheet = JT1_JT2_workbook.Sheets(2)
            temp_Exec_Time =  float(Exec_Time) * 2
            excel.Application.Run("Interface_VBA.xls!module8.JT1_JT2_Trace_Judgment",CANID_Failsafe,temp_Exec_Time,Dspace_Trace_path)            
            canid_off_time_stamp = JT1_JT2_worksheet.Cells(1,2).Value
            JT1_JT2_workbook.Save()
            JT1_JT2_workbook.Close()

        if signal_input_sys_col != None or FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2':
            print "signal_input_sys_col",signal_input_sys_col
            #Result_FailSafe_WorkBook.Close()
            excel= win32com.client.dynamic.Dispatch("Excel.Application")
            Result_FailSafe_WorkBook =  excel.Workbooks.Open(str(result_xls_path))
            Result_FailSafe_Canape_worksheet = Result_FailSafe_WorkBook.Sheets(1)
##            Address_signal = Result_FailSafe_Canape_worksheet.Columns(signal_input_sys_col).Address
##            Address_signal = Address_signal.split("$")[2]            
##            Result_Failsafe_Range = Result_FailSafe_Canape_worksheet.Range(Address_signal + "3:" + Address_signal + str(number_of_rows))            
            start_jugment = False
            start_jugment_JT1_JT2 = False
            Result_FailSafe_WorkBook.Close()
            FAILSAFE_progressbar["value"] = 11 
            Result_FailSafe_WorkBook = xlrd.open_workbook(str(result_xls_path),formatting_info=True)
            Result_FailSafe_Canape_worksheet = Result_FailSafe_WorkBook.sheet_by_index(0)
            result_write_complete = False
            if (FailSafe_Category_Pre != 'JT1' or Exec_Type != 'TYPE_B'):
                print Actual_DTC_set 
                for DTC_subArray_loop in range(0,10):
                    print DTC_subArray_loop,"DTC_subArray_loop"
                    print Actual_DTC_subArray_set[DTC_subArray_loop]
                    print Actual_DTC_set[0]
                    if Actual_DTC_subArray_set[DTC_subArray_loop] == 11:
                        print Actual_DTC_set[DTC_subArray_loop]
                        Actual_DTC_value = hex(int(Actual_DTC_set[DTC_subArray_loop]))
                        break;
                print Actual_DTC_value
                print Actual_DTC_set
            
            TYPE_B_done_Currect_behivour = False
            TYPE_C_done_Currect_behivour = False
            print Number_of_xls_file
            for Number_of_xls_file_result in range(Number_of_xls_file + 1):
                print 'Come in number of xls loop'
                result_xls_path = result_file_name_without_extension[0] + str(Number_of_xls_file_result) + '.xls'
                Result_FailSafe_WorkBook = xlrd.open_workbook(str(result_xls_path),formatting_info=True)
                Result_FailSafe_Canape_worksheet = Result_FailSafe_WorkBook.sheet_by_index(0)
                number_of_rows = Result_FailSafe_Canape_worksheet.nrows
                number_of_columns = Result_FailSafe_Canape_worksheet.ncols
                for row_canape_result in range(3,number_of_rows):            
                    present_time_stamp = Result_FailSafe_Canape_worksheet.cell(row_canape_result,0).value
                    float_present_time_stamp = float(present_time_stamp)
                    if not (FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2'):                        
                        present_value = Result_FailSafe_Canape_worksheet.cell(row_canape_result,signal_input_sys_col).value
                        float_present_value = float(present_value)            
                    if ( float_present_time_stamp >= int('1') and start_jugment == False and (not (FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2'))):
                        start_jugment = True
                        previos_xls_number = Number_of_xls_file_result 
                        float_previos_time_stamp = float(present_time_stamp)
                        float_previos_value = float(present_value)
                        previos_value = present_value
                        previos_value_row = row_canape_result
                    if ((FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2') and float_present_time_stamp >= float(canid_off_time_stamp) and start_jugment_JT1_JT2 == False):
                        start_jugment_JT1_JT2 = True
                        previos_xls_number = Number_of_xls_file_result
                        previos_value = None
                        present_value = None
                        float_present_time_stamp = 0
                        float_previos_time_stamp = canid_off_time_stamp
##                        float_previos_time_stamp = float(present_time_stamp)
##                        float_previos_value = float(present_value)
##                        previos_value = present_value
                        previos_value_row = row_canape_result                                            
                    if start_jugment == True or start_jugment_JT1_JT2 == True:
##                        print 'Start Judgment'
                        if (previos_value != present_value or start_jugment_JT1_JT2 == True ):
                            time_stamp_difference = float_present_time_stamp - float_previos_time_stamp
##                            print time_stamp_difference
##                            print signal_time_period
##                            print previos_value
##                            print Input_Value
                            if ((time_stamp_difference >= float(signal_time_period) and \
                                ( float(previos_value) == float(Input_Value) or FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2' )) or start_jugment_JT1_JT2 == True):
                                FAILSAFE_progressbar["value"] = 12
                                print 'Come In Value set loop'
                                for afs_status_xls_loop in range(previos_xls_number,Number_of_xls_file+1):
                                    result_xls_path = result_file_name_without_extension[0] + str(afs_status_xls_loop) + '.xls'
                                    Result_FailSafe_WorkBook = xlrd.open_workbook(str(result_xls_path),formatting_info=True)
                                    Result_FailSafe_Canape_worksheet = Result_FailSafe_WorkBook.sheet_by_index(0)                                
                                    if afs_status_xls_loop == previos_xls_number:
                                        afs_status_check_start_row = previos_value_row
                                        afs_status_check_end_row = Result_FailSafe_Canape_worksheet.nrows
                                    elif afs_status_xls_loop == previos_xls_number:
                                        afs_status_check_start_row = 2
                                        afs_status_check_end_row = number_of_rows
                                    else:
                                        afs_status_check_start_row = 2
                                        afs_status_check_end_row = Result_FailSafe_Canape_worksheet.nrows
                                    print 'afs_status_check_start_row ',afs_status_check_start_row
                                    print 'afs_status_check_end_row ',afs_status_check_end_row
                                    afs_status_not_zero = False
                                    for stuck_1_time_period in range(afs_status_check_start_row,afs_status_check_end_row):
                                        for afs_status_col_varaible in range(0,Output_Signal_Count):
                                            afs_status_value_list[afs_status_col_varaible] = Result_FailSafe_Canape_worksheet.cell(stuck_1_time_period,afs_status_col_list[afs_status_col_varaible]).value
                                            if float(afs_status_value_list[afs_status_col_varaible]) != 0 :
                                                afs_status_not_zero = True
                                        if afs_status_not_zero == True:
                                            afs_status_list=[]
                                            afs_status_list = afs_status_value_list
                                            if (FailSafe_Category_Pre == 'JT1' or Exec_Type == 'TYPE_B' ):                                                
                                                for JT1_Row_Increment in range(stuck_1_time_period,number_of_rows):
                                                    afs_status_jt1_zero = False
                                                    for afs_status_col_varaible in range(0,Output_Signal_Count):
                                                        afs_status_JT1_value_list[afs_status_col_varaible] = Result_FailSafe_Canape_worksheet.cell(JT1_Row_Increment,afs_status_col_list[afs_status_col_varaible]).value
                                                        if float(afs_status_JT1_value_list[afs_status_col_varaible]) != 0:
                                                            afs_status_jt1_zero = True
                                                    if afs_status_jt1_zero == False:                                                    
                                                        afs_status_change_time_Stamp = Result_FailSafe_Canape_worksheet.cell(stuck_1_time_period,0).value
                                                        afs_status_change_row_number = stuck_1_time_period
                                                        afs_status_change_xls_number = afs_status_xls_loop
                                                        time_Difference_On_Off = float(afs_status_change_time_Stamp) - float_previos_time_stamp
                                                        First_screen_shot_time = float(afs_status_change_time_Stamp)
                                                        Input_screen_shot_time = float_previos_time_stamp
                                                        TYPE_B_done_Currect_behivour = True
                                                        break
                                            elif (FailSafe_Category_Pre == 'JT2' or Exec_Type == 'TYPE_C'):
                                                for JT2_Row_Increment in range(stuck_1_time_period,number_of_rows):
                                                    afs_status_jt2_not_same = False
                                                    for afs_status_col_varaible in range(0,Output_Signal_Count):                                                
                                                        afs_status_JT2_value_list[afs_status_col_varaible] = Result_FailSafe_Canape_worksheet.cell(JT2_Row_Increment,afs_status_col_list[afs_status_col_varaible]).value
                                                        if float(afs_status_JT2_value_list[afs_status_col_varaible]) != float(afs_status_value_list[afs_status_col_varaible]):
                                                            afs_status_jt2_not_same = True
                                                    if afs_status_jt2_not_same == True:                                                        
                                                        afs_status_list=afs_status_JT2_value_list
                                                        afs_status_change_time_Stamp = Result_FailSafe_Canape_worksheet.cell(JT2_Row_Increment,0).value
                                                        #temp_previos_time_stamp = Result_FailSafe_Canape_worksheet.cell(previos_value_row,0).value
                                                        time_Difference_On_Off = float(afs_status_change_time_Stamp) - float_previos_time_stamp
                                                        afs_status_change_row_number = JT2_Row_Increment
                                                        afs_status_change_xls_number = afs_status_xls_loop
                                                        First_screen_shot_time = float(afs_status_change_time_Stamp)
                                                        Input_screen_shot_time = float_previos_time_stamp
                                                        TYPE_C_done_Currect_behivour = True
                                                        break
                                            else:
                                                afs_status_change_time_Stamp = Result_FailSafe_Canape_worksheet.cell(stuck_1_time_period,0).value
                                                time_Difference_On_Off = float(afs_status_change_time_Stamp) - float_previos_time_stamp                                                                        #i made change here
                                                afs_status_change_row_number = stuck_1_time_period
                                                afs_status_change_xls_number = afs_status_xls_loop
                                                First_screen_shot_time = float(afs_status_change_time_Stamp)
                                                Input_screen_shot_time = float_previos_time_stamp
                                                print 'time_Difference_On_Off',time_Difference_On_Off
                                            for i in range(len(afs_status_list)):
                                                afs_status_value=afs_status_list[i]
                                                if not(float(afs_status_value) == 0):
                                                    afs_status_num_temp=i
                                                    afs_status_decimal_temp=afs_status_value
                                                    break
                                            afs_status_num=afs_status_num_temp
                                            afs_status_decimal=afs_status_decimal_temp
                                            binary_value_afs_status = unicodetobinary(afs_status_decimal)
                                            length_binary_value_afs_status=len(str(binary_value_afs_status))
                                            actual_bit_set=((afs_status_num * 32) + length_binary_value_afs_status)
                                            break
                                        #Afs status change from 0 to another value loop
                                #afsstatus chnage loop complete
                                if check_dependent_signal == True and afs_status_change_row_number != -1:
                                    Result_Dependent_Signal_Judgement = Dependent_Signal_Judgement(previos_value_row,previos_xls_number,Number_of_xls_file_result,result_xls_path, \
                                                                                                   Result_dependent_signal_Canape_col,result_file_name_without_extension, \
                                                                                                   Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value, \
                                                                                                   afs_status_change_row_number,afs_status_change_xls_number,Master_Result_Report_FailSafe_write_WorkBook, \
                                                                                                   count_hyperlinks,Result_Dependent_Signal_list)
                                    print Result_Dependent_Signal_Judgement,'Result_Dependent_Signal_Judgement'

                                Master_Result_Report_FailSafe_write_Worksheet = Master_Result_Report_FailSafe_write_WorkBook.Sheets(2)
                                Master_Result_Report_FailSafe_Row = Master_Result_Report_FailSafe_write_Worksheet.UsedRange.Rows.Count
                                Master_Result_Report_FailSafe_Col = Master_Result_Report_FailSafe_write_Worksheet.UsedRange.Columns.Count                                                    
                                print "actual_bit_set",actual_bit_set,"time_Difference_On_Off",time_Difference_On_Off
                                print "signal_name_inp_sys",signal_name_inp_sys
                                for test_result_row_temp in range(8,Master_Result_Report_FailSafe_Row):
                                    if Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_row_temp,2).Value == FailSafe_Category_Pre:
                                        print "writting in excel when enter in failsafe pre"
                                        for test_result_sheet_canid_row_temp in range(test_result_row_temp+1,Master_Result_Report_FailSafe_Row):
                                            if Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,2).Value == CANID_Failsafe or \
                                               (Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,2).Value == str(signal_name_inp_sys) and \
                                                same_signal_check == False )or \
                                               Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,2).Value == "ID" + CANID_Failsafe + '_' + str(signal_name_inp_sys) or \
                                               (Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,2).Value == str(signal_name_inp_sys) + '_' + Another_Signal_Name and \
                                                same_signal_check == True ):
                                                FAILSAFE_progressbar["value"] = 13 
                                                print "writting in excel "
                                                print 'Come here'
                                                return_result_row_temp = test_result_row_temp
                                                return_actual_bit_set = actual_bit_set
                                                return_previos_time_stamp = float_previos_time_stamp
                                                print float_present_time_stamp
                                                if ((FailSafe_Category_Pre == 'JT1' or Exec_Type == 'TYPE_B') and  TYPE_B_done_Currect_behivour == True):
                                                    print "Excel data written"
                                                    print actual_bit_set
                                                    print time_Difference_On_Off
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_SetTime_Cols[0]+1)).Value = time_Difference_On_Off
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_set_value_Cols[0]+1)).Value = actual_bit_set
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expt_Value_Cols[0]+1)).Value = Expected_Result
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expt_SetTime_Cols[0]+1)).Value = Exec_Time
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expected_DTC_Cols[0]+1)).Value = Expected_DTC                                            
                                                elif (FailSafe_Category_Pre == 'JT2' and TYPE_C_done_Currect_behivour == True):
                                                    print "Excel data written"
                                                    print actual_bit_set
                                                    print time_Difference_On_Off
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_SetTime_Cols[0]+1)).Value = time_Difference_On_Off
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_set_value_Cols[0]+1)).Value = actual_bit_set
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expt_Value_Cols[0]+1)).Value = Expected_Result
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expt_SetTime_Cols[0]+1)).Value = Exec_Time
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expected_DTC_Cols[0]+1)).Value = Expected_DTC
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_DTC_Cols[0]+1)).Value = Actual_DTC_value
                                                else:
                                                    print "Excel data written"
                                                    print actual_bit_set
                                                    print time_Difference_On_Off
                                                    print test_result_sheet_canid_row_temp
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_SetTime_Cols[0]+1)).Value = time_Difference_On_Off
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_set_value_Cols[0]+1)).Value = actual_bit_set
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expt_Value_Cols[0]+1)).Value = Expected_Result
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expt_SetTime_Cols[0]+1)).Value = Exec_Time                                            
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Expected_DTC_Cols[0]+1)).Value = Expected_DTC                                            
                                                    Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,int(Actual_DTC_Cols[0]+1)).Value = Actual_DTC_value                                            
                                                Master_Result_Report_FailSafe_write_WorkBook.Save()
        ##                                    Test_Result_Sheet_Failsafe.cell(test_result_row_temp,7).value = float_present_time_stamp
        ##                                    Test_Result_Sheet_Failsafe.cell(test_result_row_temp,9).value = actual_bit_set
                                                result_write_complete = True
                                                break
                                        if result_write_complete == True:
                                            print "Complete"
                                            break
                            if (not (FailSafe_Category_Pre == 'JT1' or FailSafe_Category_Pre == 'JT2')):
                                float_previos_value = float(present_value)                        
                                float_previos_time_stamp = float(present_time_stamp)
                                previos_value = present_value
                                previos_value_row = row_canape_result
                                previos_xls_number = Number_of_xls_file_result 
                        else:
                            float_previos_value = float(present_value)
                            previos_value = present_value
                        #judgment Start if loop complete here    
                    if result_write_complete == True:
                        print "Complete"
                        break
                
            #One Excel data check complete here
        FAILSAFE_result_entry["state"] = NORMAL 
        FAILSAFE_result_entry.delete(0,END)


        if test_result_sheet_canid_row_temp == -1 :
            print 'not found in result'
            formatTxtPath = dest_FailSafe_Result_Folder + '\\' + 'format.txt'
            pathTextFile = dest_FailSafe_Result_Folder + '\\' + 'sync.txt'
            formatFileWrite = open(formatTxtPath,'w')
            FormatFileString = '1 0 0'
            formatTxtvalueWrite= str(FormatFileString)
            formatFileWrite.write(formatTxtvalueWrite)
            formatFileWrite.close()
            syncFileWrite = open(pathTextFile,'w')
            sync_num = 9
            valueWrite= str(sync_num)
            syncFileWrite.write(valueWrite)    
            syncFileWrite.close()                                
            flag_start = 0
            while(flag_start == 0):
                            
                syncFileRead = open(pathTextFile,'r')
                valueRead = syncFileRead.read()
                syncFileRead.close()
                if (valueRead == '8'):
                    time.sleep(1);
                    flag_start = 1                        
            
            FAILSAFE_result_entry.insert(0, "Not Found in Result") 
        else:
            Test_Result_id_Failsafe_categary = Master_Result_Report_FailSafe_write_Worksheet.Cells(test_result_sheet_canid_row_temp,Test_result_varaint_Cols[0]+1).Value
            print "gan",Test_Result_id_Failsafe_categary
            if Test_Result_id_Failsafe_categary  == None:
                print 'not found in result'
                formatTxtPath = dest_FailSafe_Result_Folder + '\\' + 'format.txt'
                pathTextFile = dest_FailSafe_Result_Folder + '\\' + 'sync.txt'
                formatFileWrite = open(formatTxtPath,'w')
                FormatFileString = '1 0 0'
                formatTxtvalueWrite= str(FormatFileString)
                formatFileWrite.write(formatTxtvalueWrite)
                formatFileWrite.close()
                syncFileWrite = open(pathTextFile,'w')
                sync_num = 9
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)    
                syncFileWrite.close()                                
                flag_start = 0
                while(flag_start == 0):
                                
                    syncFileRead = open(pathTextFile,'r')
                    valueRead = syncFileRead.read()
                    syncFileRead.close()
                    if (valueRead == '8'):
                        time.sleep(1);
                        flag_start = 1                        
                
                FAILSAFE_result_entry.insert(0, "Not Found in Result")
            else:
                formatTxtPath = dest_FailSafe_Result_Folder + '\\' + 'format.txt'
                pathTextFile = dest_FailSafe_Result_Folder + '\\' + 'sync.txt'
                formatFileWrite = open(formatTxtPath,'w')
                if First_screen_shot_time != -150 and Input_screen_shot_time != -150 :
                    FormatFileString = '1 2 ' + str(Input_screen_shot_time) + ' ' + str(First_screen_shot_time) 
                else:
                    FormatFileString = '1 0 0'
                formatTxtvalueWrite= str(FormatFileString)
                formatFileWrite.write(formatTxtvalueWrite)
                formatFileWrite.close()
                syncFileWrite = open(pathTextFile,'w')
                sync_num = 9
                valueWrite= str(sync_num)
                syncFileWrite.write(valueWrite)    
                syncFileWrite.close()                                
                flag_start = 0
                while(flag_start == 0):
                                
                    syncFileRead = open(pathTextFile,'r')
                    valueRead = syncFileRead.read()
                    syncFileRead.close()
                    if (valueRead == '8'):
                        time.sleep(1);
                        flag_start = 1                                    
                print count_hyperlinks
                if Test_Result_id_Failsafe_categary == "CA" or Test_Result_id_Failsafe_categary == "OK":
                    Master_Result_Report_FailSafe_Screenshot_worksheet = Master_Result_Report_FailSafe_write_WorkBook.Sheets(3)
                    if check_dependent_signal == True and (Result_Dependent_Signal_Judgement == "CA" and Test_Result_id_Failsafe_categary == "OK"):                        
                        Test_Result_id_Failsafe_categary = "CA"                        
                    Master_Result_Report_FailSafe_Screenshot_worksheet = Master_Result_Report_FailSafe_write_WorkBook.Sheets(3)
                    Master_Result_Report_FailSafe_Demo = Master_Result_Report_FailSafe_write_WorkBook.Sheets(4)
                    Master_Result_Report_FailSafe_Demo.Range("B4:K6").Copy()
                    reportTableTargetAdress = "B" + str(4 + (count_hyperlinks * 80)) + ":K" + str(4 + (count_hyperlinks * 80));                    
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Range(reportTableTargetAdress).PasteSpecial()                    
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Range("B" + str(4 + (count_hyperlinks * 80)) + ":"  + "E" + str(4 + (count_hyperlinks * 80))).Merge()
                    dblLeft = Master_Result_Report_FailSafe_Screenshot_worksheet.Range("B" + str(9 + count_hyperlinks + (count_hyperlinks * 80))).Left
                    dblTop = Master_Result_Report_FailSafe_Screenshot_worksheet.Range("B" + str(9 + count_hyperlinks + (count_hyperlinks * 80))).Top
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Shapes.AddPicture(Screen_shot_path,False,True,dblLeft,dblTop,-1,-1)
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + (count_hyperlinks * 80),2).Value = CANID_Failsafe + "_" + FailSafe_Category_Pre
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks * 80),2).Value = Exec_Time
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks * 80),5).Value = time_Difference_On_Off
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks * 80),6).Value = Expected_Result
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks * 80),7).Value = actual_bit_set
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks * 80),8).Value = Test_Result_id_Failsafe_categary
                    #Master_Result_Report_FailSafe_write_WorkBook.Sheets(3).Range("A1:AZ10").Columns.Autofit
                    Hyperlink_address_string = '=HYPERLINK("#Screenshots!B' + str(4 + (count_hyperlinks * 80)) + '","' + Test_Result_id_Failsafe_categary + '")'
                    print Hyperlink_address_string
                    excel.Worksheets(2).Cells(test_result_sheet_canid_row_temp,Test_result_varaint_Cols[0]+1).Value = Hyperlink_address_string
                    excel.Worksheets(1).Activate()
                    count_hyperlinks = count_hyperlinks + 1
                if Test_Result_id_Failsafe_categary == "CA":
                    CA_Report_xls_file = CA_Dest_Folder + '\\' + 'Inspection_Caution_Advisory_Report_V1.xls'
                    CA_Report_TestProcedure_file = CA_Test_Procedure_Folder + '\\' + 'Failsafe.xls'  #gana
                    print CA_Report_TestProcedure_file
                    CA_Report_Workbook = excel.Workbooks.Open(str(CA_Report_xls_file))
                    CA_Report_Test_Procedure_Workbook = excel.Workbooks.Open(str(CA_Report_TestProcedure_file))
                    CA_sheet_number = count_hyperlinks_CA + 1
                    CA_Report_Workbook.Worksheets(CA_sheet_number).Copy(Before=CA_Report_Workbook.Worksheets(CA_sheet_number))
                    CA_sheet_Name = 'CA_' + str(count_hyperlinks_CA)
                    CA_Report_Workbook.Worksheets(CA_sheet_number).Name = CA_sheet_Name
                    CA_Hyperlink_address_string = '=HYPERLINK("#CA_' + str(count_hyperlinks_CA) + '!A1","' + CA_sheet_Name + '")'
                    CA_Report_Workbook.Worksheets(1).Cells( 8 + count_hyperlinks_CA ,2).Value = CA_Hyperlink_address_string
                    CA_Report_Workbook.Worksheets(1).Range(CA_Report_Workbook.Worksheets(1).Cells(8 + count_hyperlinks_CA ,2), \
                                                                 CA_Report_Workbook.Worksheets(1).Cells(8 + count_hyperlinks_CA ,4)).Borders.Weight = 3
                    reportTableTargetAdress = "B" + str(4 + (count_hyperlinks * 80)) + ":K" + str(7 + (count_hyperlinks * 80))                    
                    Master_Result_Report_FailSafe_Screenshot_worksheet.Range(reportTableTargetAdress).Copy()
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Range("B10:K13").PasteSpecial()
                    CA_Report_Workbook.Worksheets(1).Range("I9").Copy()                    
                    CA_Result_Color_Range = "D" + str(8 + count_hyperlinks_CA)                
                    CA_Report_Workbook.Worksheets(1).Range(CA_Result_Color_Range).PasteSpecial()
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Shapes.AddPicture(Screen_shot_path,False,True,33,180,-1,-1)
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(4,3).Value = CANID_Failsafe + "_" + FailSafe_Category_Pre
                    if (Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks_CA * 80),6).Value == Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(6 + (count_hyperlinks_CA * 80),7).Value):
                        CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(6,3).Value = "Output deviations occured because of value"
                        CA_Report_Workbook.Worksheets(1).Cells( 8 + count_hyperlinks_CA ,3).Value = "Output deviations occured because of value"
                    else:
                        CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(6,3).Value = "Output deviations occured at timestamps"
                        CA_Report_Workbook.Worksheets(1).Cells( 8 + count_hyperlinks_CA ,3).Value =   "Output deviations occured at timestamps"
                    col_test_procedure = 3
                    Test_procedure_data = ''
                    print 'Failsafe_Signals',FailSafe_Category_Pre
                    for row_test_procedure in range (1,CA_Report_Test_Procedure_Workbook.Worksheets(1).UsedRange.Rows.Count):
                        if((CA_Report_Test_Procedure_Workbook.Worksheets(1).Cells(row_test_procedure,col_test_procedure).Value) == FailSafe_Category_Pre):
                            print CA_Report_Test_Procedure_Workbook.Worksheets(1).Cells(row_test_procedure,col_test_procedure).Value
                            Test_procedure_data = CA_Report_Test_Procedure_Workbook.Worksheets(1).Cells(row_test_procedure,(col_test_procedure + 1)).Value
                            break;
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(5,3).Value = Test_procedure_data
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(2,3).Value = Mot_file_name
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(3,3).Value = Hils_model_name
                    CA_Hyperlink_address_string = '=HYPERLINK("#CA_' + str(count_hyperlinks_CA) + '!A10","Click_Here")'
                    CA_Report_Workbook.Worksheets(count_hyperlinks_CA + 1).Cells(7,3).Value = CA_Hyperlink_address_string                    
                    CA_Report_Workbook.Worksheets(1).Activate()
                    CA_Report_Workbook.Save()
                    CA_Report_Test_Procedure_Workbook.Worksheets(1).Activate()
                    CA_Report_Test_Procedure_Workbook.Save()
                    CA_Report_Workbook.Close()
                    CA_Report_Test_Procedure_Workbook.Close()
                    count_hyperlinks_CA = count_hyperlinks_CA + 1
                FAILSAFE_result_entry.insert(0,Test_Result_id_Failsafe_categary)
                FailSafe_Display_Tree.item(curItem_FailSafe, text = CANID_Failsafe, values = Test_Result_id_Failsafe_categary )                ## Need to chnage result
        Master_Result_Report_FailSafe_write_WorkBook.Save()
        Master_Result_Report_FailSafe_write_WorkBook.Close()
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,                                  # Creation of log file
                        format='%(asctime)s - %(levelname)s - %(message)s')             
        logging.info('Judgment Completed')                                                                 # Logging info in the log file
        FAILSAFE_result_entry["state"] = DISABLED        
        time.sleep(3)                
        FAILSAFE_result_entry["state"] = NORMAL 
        FAILSAFE_result_entry.delete(0,END)
        return count_hyperlinks,count_hyperlinks_CA
                
    def unicodetobinary(unicode_value):
        float_value = float(unicode_value)
        integer_value = int(float_value)
        temp = 1
        binary_value = 0
        while(integer_value != 0):
            reminder = integer_value % 2
            integer_value = integer_value / 2
            binary_value = binary_value  + (reminder*temp)
            temp = temp * 10
        return binary_value
    
    def Dependent_Signal_Judgement(previos_value_row,previos_xls_number,Number_of_xls_file_result,result_xls_path,Result_dependent_signal_Canape_col,result_file_name_without_extension, \
                                   Result_Dependent_Signal_Comparison_list_value,Result_Dependent_Signal_list_value,afs_status_change_row_number,afs_status_change_xls_number, \
                                   Master_Result_Report_FailSafe_write_WorkBook,count_hyperlinks,Result_Dependent_Signal_list):
        #*************************************************************Multiple signal judgement ganpi***************************************************************************#
        Dependent_signal_check = None
        Absolute_value = 0
        count_multiple_result_dependent_signal = 0
        print Result_dependent_signal_Canape_col
        final_depedent_signal_status_value = range(len(Result_dependent_signal_Canape_col));
        found_multiple_result_dependent_signal = 0
        depedent_signal_status_value = range(len(Result_dependent_signal_Canape_col));
        Master_Result_Report_FailSafe_Screenshot_worksheet = Master_Result_Report_FailSafe_write_WorkBook.Sheets(3)
        Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + (count_hyperlinks * 80),26).Value = "Signal Name"
        Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + (count_hyperlinks * 80),27).Value = "Expected Value"
        Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + (count_hyperlinks * 80),28).Value = "Actual Value"
        Master_Result_Report_FailSafe_Screenshot_worksheet.Range(Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + (count_hyperlinks * 80),26), \
                                                                 Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + (count_hyperlinks * 80),28)).Borders.Weight = 3
        Master_Result_Report_FailSafe_Screenshot_worksheet.Range(Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (count_hyperlinks * 80),26), \
                                                                 Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(4 + len(Result_dependent_signal_Canape_col) + (count_hyperlinks * 80),28)).Borders.Weight = 2
        result_xls_path_depedent_signal = result_file_name_without_extension[0] + str(afs_status_change_xls_number) + '.xls'
        Result_FailSafe_WorkBook = xlrd.open_workbook(str(result_xls_path_depedent_signal),formatting_info=True)
        Result_FailSafe_Canape_worksheet = Result_FailSafe_WorkBook.sheet_by_index(0)                                
        for col_check_dept_signal in range(len(Result_dependent_signal_Canape_col)):
            if col_check_dept_signal != 0:
                if '-' in Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal - 1]:
                    print 'not to check'
                    continue
            if '-' in Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal]:
                print afs_status_change_row_number,Result_dependent_signal_Canape_col[col_check_dept_signal],"start","end","first"
                print depedent_signal_status_value[col_check_dept_signal] 
                depedent_signal_status_value[col_check_dept_signal] = Result_FailSafe_Canape_worksheet.cell(afs_status_change_row_number,Result_dependent_signal_Canape_col[col_check_dept_signal]).value
                depedent_signal_status_value[col_check_dept_signal + 1] = Result_FailSafe_Canape_worksheet.cell(afs_status_change_row_number,Result_dependent_signal_Canape_col[col_check_dept_signal + 1 ]).value
                Absolute_value = float(depedent_signal_status_value[col_check_dept_signal]) - float(depedent_signal_status_value[col_check_dept_signal + 1])
                Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (col_check_dept_signal) + (count_hyperlinks * 80),26).Value = Result_Dependent_Signal_list[col_check_dept_signal] \
                                                                                                                                           + '-' + Result_Dependent_Signal_list[col_check_dept_signal + 1]
                Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (col_check_dept_signal) + (count_hyperlinks * 80),27).Value = Result_Dependent_Signal_list_value[col_check_dept_signal]
                Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (col_check_dept_signal) + (count_hyperlinks * 80),28).Value = abs(Absolute_value)
                print abs(Absolute_value)
                print float(depedent_signal_status_value[col_check_dept_signal])
                print float(depedent_signal_status_value[col_check_dept_signal + 1])
                if Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal+1] == ">":
                    print 'Coming In Greater'
                    if (abs(float(Absolute_value)) > float(Result_Dependent_Signal_list_value[col_check_dept_signal])):                    
                        print "checking muliple depedent signal:",depedent_signal_status_value[col_check_dept_signal]
                        count_multiple_result_dependent_signal = count_multiple_result_dependent_signal + 2
                elif Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal+1] == "<":
                    print 'Coming In Less Then'
                    if (abs(float(Absolute_value)) < Result_Dependent_Signal_list_value[col_check_dept_signal]):
                        print "checking muliple depedent signal:",depedent_signal_status_value[col_check_dept_signal]
                        count_multiple_result_dependent_signal = count_multiple_result_dependent_signal + 2
                else:
                    print 'Coming In Equal'
                    if (abs(float(Absolute_value)) == Result_Dependent_Signal_list_value[col_check_dept_signal]):
                        print "checking muliple depedent signal:",depedent_signal_status_value[col_check_dept_signal]
                        count_multiple_result_dependent_signal = count_multiple_result_dependent_signal + 2                
            else:
                print afs_status_change_row_number,Result_dependent_signal_Canape_col[col_check_dept_signal],"start","end"
                print depedent_signal_status_value[col_check_dept_signal] 
                depedent_signal_status_value[col_check_dept_signal] = Result_FailSafe_Canape_worksheet.cell(afs_status_change_row_number,Result_dependent_signal_Canape_col[col_check_dept_signal]).value
                Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (col_check_dept_signal) + (count_hyperlinks * 80),26).Value = Result_Dependent_Signal_list[col_check_dept_signal]
                Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (col_check_dept_signal) + (count_hyperlinks * 80),27).Value = Result_Dependent_Signal_list_value[col_check_dept_signal]
                Master_Result_Report_FailSafe_Screenshot_worksheet.Cells(5 + (col_check_dept_signal) + (count_hyperlinks * 80),28).Value = depedent_signal_status_value[col_check_dept_signal]
                print str(depedent_signal_status_value[col_check_dept_signal]) + str(Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal]) +  str(Result_Dependent_Signal_list_value[col_check_dept_signal])
                if Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal] == ">":
                    print 'Coming In Greater'
                    if (float(depedent_signal_status_value[col_check_dept_signal]) > float(Result_Dependent_Signal_list_value[col_check_dept_signal])):                    
                        print "checking muliple depedent signal:",depedent_signal_status_value[col_check_dept_signal]
                        count_multiple_result_dependent_signal = count_multiple_result_dependent_signal + 1
                elif Result_Dependent_Signal_Comparison_list_value[col_check_dept_signal] == "<":
                    print 'Coming In Less Then'
                    if (float(depedent_signal_status_value[col_check_dept_signal]) < float(Result_Dependent_Signal_list_value[col_check_dept_signal])):
                        print "checking muliple depedent signal:",depedent_signal_status_value[col_check_dept_signal]
                        count_multiple_result_dependent_signal = count_multiple_result_dependent_signal + 1
                else:
                    print 'Coming In Equal'
                    if (float(depedent_signal_status_value[col_check_dept_signal]) == float(Result_Dependent_Signal_list_value[col_check_dept_signal])):
                        print "checking muliple depedent signal:",depedent_signal_status_value[col_check_dept_signal]
                        count_multiple_result_dependent_signal = count_multiple_result_dependent_signal + 1
        Master_Result_Report_FailSafe_write_WorkBook.Save()
        print count_multiple_result_dependent_signal,'count_multiple_result_dependent_signal'
        if (count_multiple_result_dependent_signal == len(Result_dependent_signal_Canape_col)):
            Dependent_signal_check = "OK"
        else:
            Dependent_signal_check = "CA"
        return Dependent_signal_check


    def ADAS_HILS_FAILSAFE_AUTOMATION (SignalData,SigInfo,sig_data_sheet ,myAppl):
        try:
            print "ADAS_HILS_FAILSAFE_AUTOMATION"
            wait_confirm_val = 100000
    ##        myAppl.Variable(Power_Supply_path).Write(1)
    ##        time.sleep(.5)
    ##        myAppl.Variable("simState").Write(0)                                                                    # 'Reset' Simstate
    ##        time.sleep(.5)
    ##        myAppl.Variable("simState").Write(2)                                                                    # 'Set' Simstate
    ##        time.sleep(2)
            
            
          
            for m in range(len(SignalData)):                    # Loop for execution type "exec_start_end" and "exec_cont" 
                Signal_Data = []
                Signal_Data = SignalData[m]
               
                test_sig_type = Signal_Data[0]
                if test_sig_type in ['0']:
                    sig_name =  Signal_Data[1]                                                                      # Collect signal name
                    sig_path = Signal_Data[2]                                                                       # Collect signal path  
                    sig_delay = float(Signal_Data[3])        # Collect the delay value . This delay will be executed after desired value for the signal is set 
                    sig_val = Signal_Data[4]                 # Collect the various values to be set for a particular signal. Note that 'sig_val' is a 'list' 

                    for n in range(len(sig_val)):                        
                        myAppl.Variable(sig_path).Write(sig_val[n])
                    
                        if sig_delay < 0:    # If delay specified is less than zero, this loop will confirm if the desired value is being set to the signal. Else specified delay is executed.
                            temp_count = 0 
                            while temp_count < wait_confirm_val:
                                temp_count = temp_count + 1                                
                                temp_val = None
                                temp_val = myAppl.Variable(sig_path).Read()                                         # Value written to the signal is read
                                time.sleep(0.5)
                                if temp_val == sig_val[n]:                                                          # After confirmation loop ends
                                    break
                        else:
                            time.sleep(float(sig_delay))
                    
                elif test_sig_type in ['1']:                    
                    sig_name = Signal_Data[1]                                                                       # Collect the name of the signal to which value is to be set        
                    sig_path = Signal_Data[2]                                                                       # Collect the path of the above signal
                    sig_val = Signal_Data[3]                                                                        # Collect the value to be set to the above signal
                    dep_var_name = Signal_Data[4]                                                                   # Collect the name of the dependent variable.  
                    dep_var_path = Signal_Data[5]                                                                   # Collect the path of the dependent variable
                    dep_var_cond = Signal_Data[6]                                                                   # Collect the dependency condition
                    dep_var_val = float(Signal_Data[7])                                                             # Collect the value of dependent variable 

                                                                                                                    
                    if dep_var_cond in ['>']:                                                                       # Execute this loop if dependency condition is "greater than"
                        temp_count = 0
                        

                        while temp_count < wait_confirm_val: 
                           
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            print "temp_val",temp_val
                            print "dep_var_val",dep_var_val
                            if temp_val > dep_var_val:                                                              # After the condition is met value is written to the signal
                                print "temp_val",temp_val
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Write the desired value to the signal only if dependent variable meets the condition with specified dependent value
                                break

                                                            
                    if dep_var_cond in ['=']:                                                                       # Execute this loop if dependency condition is "equal to" 
                        temp_count = 0
                        
                        while temp_count < wait_confirm_val:
                            
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))                              # Value of the dependent variable is read


                            if temp_val == dep_var_val:                                                             # After the condition is met value is written to the signal
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Write the desired value to the signal only if dependent variable meets the condition with specified dependent value
                                break

                                             
                        
                    if dep_var_cond in ['<']:                                                                       # Execute this loop if dependency condition is "less than" 
                        temp_count = 0
                       
                        while temp_count < wait_confirm_val:
                           
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))                              # Value of the dependent variable is read

                            if temp_val < dep_var_val:                                                              # After the condition is met value is written to the signal
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Write the desired value to the signal only if dependent variable meets the condition with specified dependent value
                                break
                            
                elif test_sig_type in ['2']:
                    sig_name = Signal_Data[1]
                    sig_delay = float(Signal_Data[2])
                    print "\n waiting for delay = " + str(sig_delay)
                    time.sleep(sig_delay)

           
                
                elif test_sig_type in ['3']:
                    sig_name = Signal_Data[1]                                                                       # Collect the name of the signal to which value is to be set        
                    sig_path = Signal_Data[2]                                                                       # Collect the path of the above signal
                    sig_val = Signal_Data[3]                                                                        # Collect the value to be set to the above signal
                    dep_var_name = Signal_Data[4]                                                                   # Collect the name of the dependent variable.  
                    dep_var_path = Signal_Data[5]                                                                   # Collect the path of the dependent variable
                    dep_var_cond = Signal_Data[6]                                                                   # Collect the dependency condition
                    dep_var_val = float(Signal_Data[7])                                                             # Collect the value of dependent variable
              

                                                                 
                
                    if dep_var_cond in ['=']:                                                                       # Execute this loop if dependency condition is "equal to" 
                        temp_count = 0
                        while temp_count < wait_confirm_val:
                            time.sleep(0.5)
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            temp_count = temp_count + 1                           
                            
                            if temp_val == dep_var_val:                                                             # After the condition is met....value is written to the signal 
                               break
                            else:                                
                                myAppl.Variable(sig_path).Write(0)                                                  # Writing the value '0' is push button feature
                                time.sleep(0.1)
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Writing the desired value to the signal
                                time.sleep(0.1) 
                                myAppl.Variable(sig_path).Write(0)

                                                   
              
                
                    if dep_var_cond in ['>']:                                                                       # Execute this loop if dependency condition is "greater than" 
                        temp_count = 0
                        while temp_count < wait_confirm_val:
                            time.sleep(0.5)
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            temp_count = temp_count + 1                           
                            
                            if temp_val > dep_var_val:                                                              # After the condition is met....value is written to the signal 
                               break
                            else:                                
                                myAppl.Variable(sig_path).Write(0)                                                  # Writing the value '0' is push button feature
                                time.sleep(0.1)
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Writing the desired value to the signal
                                time.sleep(0.1) 
                                myAppl.Variable(sig_path).Write(0)

             
                
                    if dep_var_cond in ['<']:                                                                       # Execute this loop if dependency condition is "less than" ##
                        temp_count = 0
                        while temp_count < wait_confirm_val:
                            time.sleep(0.5)
                            temp_val = myAppl.Variable(dep_var_path).Read()                                         # Value of the dependent variable is read
                            temp_count = temp_count + 1                           
                            
                            if temp_val < dep_var_val:                                                              # After the condition is met....value is written to the signal 
                               break
                            else:                                
                                myAppl.Variable(sig_path).Write(0)                                                  # Writing the value '0' is push button feature
                                time.sleep(0.1)
                                myAppl.Variable(sig_path).Write(sig_val)                                            # Writing the desired value to the signal
                                time.sleep(0.1) 
                                myAppl.Variable(sig_path).Write(0)


            
                                
                elif test_sig_type in ['4']:                    
                    sig_name = Signal_Data[1]                    
                    dep_var_name = Signal_Data[2]
                    dep_var_path = Signal_Data[3]
                    dep_var_cond = Signal_Data[4]
                    dep_var_val = float(Signal_Data[5])
                
                    if dep_var_cond in ['>']:
                        
                        while(1): 
                            
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))
                            time.sleep(0.2)
                            
                            if temp_val > dep_var_val:
                               break
                        
                    elif dep_var_cond in ['<']:
                        
                        while temp_count < wait_confirm_val:
                            temp_count = temp_count + 1
                            temp_val = -1
                            temp_val = myAppl.Variable(dep_var_path).Read()
                            time.sleep(0.2)                          
                            if temp_val < dep_var_val:
                               # myAppl.Variable(sig_path).Write(sig_val)
                                break
                        
                    elif dep_var_cond in ['=']:

                        while(1): 
                            
                            temp_val = -1
                            temp_val =float(str(myAppl.Variable(dep_var_path).Read()))
                            time.sleep(0.2)
                            if temp_val == dep_var_val:
                                break
                elif test_sig_type in ['5']:                    
                                    
                    dep_var_name = Signal_Data[2]
                    sig_path = Signal_Data[3]
                    dep_var_val = float(Signal_Data[5])

                    dep_var_name = dep_var_name + ";"
                    temp_sig_path = sig_path
                    flag = 0
                    i = 0
                    old_Rx_status = 9
                    Tx_status = 16
                    while ((len(dep_var_name)) >= 2):
                        sig_path = temp_sig_path
                        
                        if i == len(dep_var_name):
                             break;   
                        if i == 0:
                           sig_path = sig_path + "1/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                           
                        elif i == 1:
                           sig_path = sig_path + "2/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                           
                        elif i == 2:
                           sig_path = sig_path + "3/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 3:
                           sig_path = sig_path + "4/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 4:
                           sig_path = sig_path + "5/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 5:
                           sig_path = sig_path + "6/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 6:
                           sig_path = sig_path + "7/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                        elif i == 7:
                           sig_path = sig_path + "8/Value"
                           myAppl.Variable(sig_path).Write(ord(dep_var_name[i]))
                           temp_dep_var_name = ''
                           for j in range(i+1,len(dep_var_name)):
                               temp_dep_var_name = temp_dep_var_name + dep_var_name[j]
                           dep_var_name = temp_dep_var_name
                           i = -1
                        
                           time.sleep(1)
                           while(1):

                                  Rx_status = myAppl.Variable(Test_Automation_path)\
                                              .Read()
                                  time.sleep(0.01)
                                
                                  if Rx_status > old_Rx_status:
                                       old_Rx_status = Rx_status
                                       myAppl.Variable(Start_CANape_path)\
                                       .Write(Tx_status)
                                       Tx_status = Tx_status + 1
                                       break;
                              

                time.sleep(0.5)
                sig_path = None                                                                                     # Set default values to all the signals present in sheet2 of excel sheet
                sig_name = None
                sig_value = None
                sig_reset = None

                time.sleep(2)
                #myAppl.Variable(Power_Supply_path).Write(129)            


        except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.exception('Signal_paths are wrong')
            
    def Message_Counter_Testing() :        
       # SwitchTab()
        global Var_Val, myAppl,DispatchSheet,Message_Counter_Variant_Value
        overall_progressbar=0
        global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DIAG_CMD_NO_path_3,DTC_string_temp,DTC_string_1_temp,DTC_string_path_temp,Read_vehicle_speed_path
          
       # nb.select(Message_counter_background_frame)
        t0 = time.time()
        t1 = time.time()
        
 
        
        Message_counter_overall_progressbar["value"] = overall_progressbar
        
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(message)s')
        logging.info('##############  Start of Message Counter  Testing  ##############')
        Message_Counter_Dict = OrderedDict()                    # This makes dictionary required for Message Counter Tree
        Message_Counter_Dict[VehicleName]= OrderedDict()
        Message_Counter_Dict[VehicleName][RegionName]= OrderedDict()
        Message_Counter_Dict[VehicleName][RegionName][PartNo]= OrderedDict()
        Active_Test = 'MESSAGE_COUNTER'
        Message_Counter_Dict[VehicleName][RegionName][PartNo][Active_Test]=OrderedDict()
        
        Message_Counter_Result_Folder=[]     # This contains the path of result folders created individually for all the variants for Message Counter result
        Message_Counter_CANID_list = []      #This contains list of all ENABLED Message Counter CANID extracted from the dispatch sheet      
        COUNT_YES = []                       #This stores the number number of Enabled Message Counter CANID for each variant. i.e. No of CANID with 'Y' in front of them        
  
        
        
        DIMPSheet_Message_Counter_CANID_List = WorkBook.sheet_by_index(1)          #This opens the SECOND Sheet of DISPATCH Sheet i.e.Sheet containing list of active Message Counter CANID
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
        logging.info('Dispatch Sheet for Message Counter Loaded')		
        DIMPSheet_Message_Counter_CANID_List_Col = DIMPSheet_Message_Counter_CANID_List.ncols     
        DIMPSheet_Message_Counter_CANID_List_Row = DIMPSheet_Message_Counter_CANID_List.nrows
        Message_Counter_CANID_row=[]     #List used to store row number of "Y" in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET).       Used to insert result later in code 
        Message_Counter_CANID_column=[]  #List used to store column number of "Y" in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET).    Used to insert result later in code

        
       
        Message_counter_overall_progressbar["value"]=overall_progressbar
        try:
            for k in range(0,len(Message_Counter_Variant_Value)):
                
                counter = 0   
                Message_Counter_Dict[VehicleName][RegionName][PartNo][Active_Test][Message_Counter_Variant_Value[k]]=OrderedDict()
                Message_Counter_Result_Folder.append(VehicleNameFolder[k]+"\\07_Message_Counter")              # This is Destination folder to store copy of MESSAGE_COUNTER CANape configuration
                dest_Message_Counter_Result_Folder = Message_Counter_Result_Folder[k]                          
                distutils.dir_util.copy_tree(Master_CANape_Message_Counter_Path,dest_Message_Counter_Result_Folder)
                #copy_folder(Master_CANape_Message_Counter_Path,dest_Message_Counter_Result_Folder)        # This function copies files and folders from MASTER CANAPE CONFIG 
                if (k < 10):
                    Var_str =  'Variant_0' +  str(k +1)    #This is used to obtain Variant number in format "Variant_0X" eg. Variant_01 Variant_02 etc
                else:
                    Var_str =  'Variant_' +  str(k +1) 
                Var_Row = 0     #Store row number of Variant in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET) 
                Var_col = 0     #Store column number of Variant in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)
                CANID_row=0     #Store row number of the CANID in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)
                CANID_col=0     #Store column number of the CANID in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)
                
                
                try:
                    for i in range (0,  DIMPSheet_Message_Counter_CANID_List_Row):             # Loop for traversing through the EXCEL sheet
                        for j in range(0,DIMPSheet_Message_Counter_CANID_List_Col):
                            if DIMPSheet_Message_Counter_CANID_List.cell(i,j).value == Var_str:         #This finds row and column of the particular variant
                                Var_col =  j
                                Var_Row = i
                                
                            if DIMPSheet_Message_Counter_CANID_List.cell(i,j).value == 'CANID':        #This finds row and column of CANIDs in MSG_COUNTER_DETAIL
                                 CANID_col =  j
                                 CANID_row = i
                            else:
                                continue
                except:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                            level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('Rows and columns not found for CANID and variant')

                try:                           
                    for i in range (Var_Row + 2, DIMPSheet_Message_Counter_CANID_List_Row):      # ( Var_Row +2 ) contains the string "Y" or "N"
                       
                        if (DIMPSheet_Message_Counter_CANID_List.cell(i,Var_col).value=='Y'):
                            counter = counter +1           #This counts the number of 'Y' for a particular variant. 
                            CANID_msg_ctr=DIMPSheet_Message_Counter_CANID_List.cell(i,CANID_col).value    #This extracts the CANID form the sheet
                            Message_Counter_Dict[VehicleName][RegionName][PartNo][Active_Test][Message_Counter_Variant_Value[k]][CANID_msg_ctr]=OrderedDict()   #This adds CANID to the Message Counter Tree
                            Message_Counter_CANID_list.append(DIMPSheet_Message_Counter_CANID_List.cell(i,CANID_col).value)   #This adds CANID to a list 
                            Message_Counter_CANID_row.append(i)   #This stores the row number of all "Y" for future reference
                    Message_Counter_CANID_column.append(Var_col)  #This store the column number of all "Y" for future reference 
                    COUNT_YES.append(counter)                      #This is the final count of "Y". Each element of COUNT_YES correspons to the corresponding Variant
                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                            level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('Y value not found')   
                
            print "Message_Counter_Dict",Message_Counter_Dict
            uid_MSG_prev=uid
            Message_Counter_Tree = construct_JSON_tree(Message_Counter_Dict,frame8)    #This makes Message Counter Tree
            curItem_MSG_CTR= uid_MSG_prev+1     #Since uid doest become 0
            Message_counter_vehicle_id = Message_Counter_Tree.item(curItem_MSG_CTR, 'text')   #Extracts information from Message Counter Tree      
            Message_counter_vehicle_id_entry.insert(0, Message_counter_vehicle_id)            #Fills the space in Message Counter GUI with required information  

            curItem_MSG_CTR=curItem_MSG_CTR+3   #This is used to jump over the items not required in Tree for displaying in GUI

            counter_row=0 #Simple counter variable for incrementing 
            print "myAppl",myAppl
            for k in range(0,len(Message_Counter_Variant_Value)):
                progressbar=0
                Message_counter_progressbar["maximum"] = (COUNT_YES[k])+7
                Message_counter_overall_progressbar["maximum"] = len(Message_Counter_Variant_Value)
                Message_counter_overall_progressbar["value"] = overall_progressbar
                Message_counter_progressbar['value']=progressbar

                InterfacevalueRead_MSG=1
                
                Message_counter_variant_entry.delete(0,END)      #Clears the space in Message Counter GUI  
                Message_counter_CAN_ID_entry.delete(0,END)       #Clears the space in Message Counter GUI 
                Message_counter_result_entry.delete(0,END)       #Clears the space in Message Counter GUI   
                
                SyncPathTextFile = Message_Counter_Result_Folder[k] + "\\" + "Sync.txt" #Sync.txt is used to sync PYTHON and CANAPE (CAPL script)
                CSVFile = Message_Counter_Result_Folder[k] + "\\" + "CANape.txt"        #CANape.txt stores data converted from MDF file. Used for judgement

                
                Var_str =  'Variant_0' +  str(k +1)  #This is used to obtain Variant number in format "Variant_0X" eg. Variant_01 Variant_02 etc
                
                curItem_MSG_CTR = curItem_MSG_CTR + 1 
                Message_Counter_Tree.selection_set(curItem_MSG_CTR)   #Highlighting the item in the Tree
       
                Variant_Name = Message_Counter_Tree.item(curItem_MSG_CTR, 'text')  
                Message_counter_variant_entry.insert(0, Variant_Name)
                
                Write_Var = Message_Counter_Variant_Value[k]
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('Variant %s under Message Counter testing',Write_Var)			

                try:
                    myAppl.Variable(Power_Supply_path).Write(1)   # Switch on Power Supply to Write Variant Code     
                            ##Instrumentation().ActiveLayout.Normalize()
                    Variant_write(myAppl,Write_Var)
                except Exception, e:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                            level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('variant code not written error in power supply path')    

                progressbar=progressbar+1
                Message_counter_progressbar['value']=progressbar                     

             
                valueRead_MSG = 1                                    
                syncFileWriteMSG = open(SyncPathTextFile,'w')
                sync_num = 7                                    #Write 7 in Sync.txt to erase any garbage 
                valueWrite= str(sync_num)
                syncFileWriteMSG.write(valueWrite)
                syncFileWriteMSG.close()
        
                myAppl.Variable(Power_Supply_path).Write(1)

                try:
                    Start_CANape(Message_Counter_Result_Folder[k])    #Start Canape for Message Counter
                except:
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                            level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('CANape not started')    
    
                
                sync_num = 9                                      #Write '9' in Sync.txt to tell CAPle to start working
                syncFileWriteMSG = open(SyncPathTextFile,'w')
                valueWrite= str(sync_num)
                syncFileWriteMSG.write(valueWrite)
                syncFileWriteMSG.close()
                
                while(valueRead_MSG != '8'):                    #Wait for CAPle to tell python that it has finished CANape and other processing work like  MDF conversion
                    syncFileRead = open(SyncPathTextFile,'r')   #CAPle will write '8' in Sync.txt to inform python that it has finished its work
                    valueRead_MSG = syncFileRead.read()
                    syncFileRead.close()
                    time.sleep(1);
                    
                progressbar=progressbar+3
                Message_counter_progressbar["value"]=progressbar
                
                myAppl.Variable(Power_Supply_path).Write(129)   #CANape work done. Switch off power supply

                print "execution done "
                xlapp = win32com.client.dynamic.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet 

                if os.path.exists(str(CAN_MSG_DATA_PATH)):

                    xlapp.Workbooks.Open(Filename=str(CAN_MSG_DATA_PATH), ReadOnly=1)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                            level=logging.INFO,
                                            format='%(message)s')
                    logging.info('Judgement for Message Counter started')
                
                    print "Result_DispatchSheet_Path" , Result_DispatchSheet_Path
                    ##print "CSVFile" , CSVFile
                    ##print "InterfaceTextFile" , InterfaceTextFile
                    xlapp.Application.Run("Message_Counter_Judgement_Sheet.xls!module2.number_of_varaint",Var_str,str(Result_DispatchSheet_Path),str(CSVFile),str(InterfaceTextFile))
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                            level=logging.INFO,
                                            format='%(message)s')
                    logging.info('Judgement for Message Counter finished')

                    xlapp.Application.Quit() # Comment this out if your excel script closes
                    shutil.copy(Message_Counter_Report_path,Message_Counter_Result_Folder[k])
                    Result_Message_Counter_Report_path = Message_Counter_Result_Folder[k] + "\\"  + "Message_Counter_Report.xlsx"
                    xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
                    xlapp.Workbooks.Open(Filename=str(interface_sheet_path), ReadOnly=1)
                    xlapp.Application.Run("Interface_VBA.xls!module9.make_result_xls",str(DispatchSheet),Var_str,Result_Message_Counter_Report_path,str(Result_DispatchSheet_Path),"MSG_COUNTER","MSG_COUNTER")
                    xlapp.Application.Quit() # Comment this out if your excel script closes
                    

                # When macro in Message_Counter_Judgement_Sheet has generated the result for Message Counter in Dispatch Sheet macro writes '8' in Interface.txt file.close
                #If 8 is not written Python sleeps and waits for Macro to write 8
                progressbar=progressbar+3
                Message_counter_progressbar["value"]=progressbar
            
                while(InterfacevalueRead_MSG != '8'):           
                    time.sleep(1)                  
                    InterfacesyncFileRead = open(InterfaceTextFile,'r')
                    InterfacevalueRead_MSG = InterfacesyncFileRead.read(1)
                    InterfacesyncFileRead.close()
                    
                os.remove(InterfaceTextFile) #This deletes the Interface.txt file
                
                Dispatch_Sheet_Result_WorkBook =xlrd.open_workbook(str(Result_DispatchSheet_Path),formatting_info=True) #This opens Result DISPATCH SHEET Workbook
                Message_Counter_Result_Sheet=Dispatch_Sheet_Result_WorkBook.sheet_by_index(1)   #The Message Counter Sheet in DISPATCH SHEET WORKBOOK
      
                for t in range(0,COUNT_YES[k]):
                    Message_counter_result_entry.delete(0,END)   #Clears the space in Message Counter GUI
                    Message_counter_CAN_ID_entry.delete(0,END)   #Clears the space in Message Counter GUI
                    curItem_MSG_CTR=curItem_MSG_CTR+1            #
                    Message_Counter_Tree.selection_set(curItem_MSG_CTR)
                    CANID_MSG_CTR = Message_Counter_Tree.item(curItem_MSG_CTR, 'text')
                    Message_counter_CAN_ID_entry.insert(0, CANID_MSG_CTR )                
     ##              while (Message_Counter_Result_Sheet.cell(Message_Counter_CANID_row[counter_row],Message_Counter_CANID_column[counter_col]).value == 'Y'):
    ##                  time.sleep(1)
                    #To display Test Case result on Message Counter GUI. Logic is as follows
                    #Message_Counter_CANID_row[counter_row]   Message_Counter_CANID_row is a list obtained earlier containing row number of "Y".
                    #Message_Counter_CANID_column  is a list obtained earlier containing column number of "Y".
                    #"Y" is replaced by "OK"  or "CA" or "RESULT NOT FOUND"    . So row number obtained earlier is utilized      
                    Message_Counter_Tree.item(curItem_MSG_CTR, text = CANID_MSG_CTR, values = Message_Counter_Result_Sheet.cell(Message_Counter_CANID_row[counter_row],Message_Counter_CANID_column[k]).value)

                    if Message_Counter_Result_Sheet.cell(Message_Counter_CANID_row[counter_row],Message_Counter_CANID_column[k]).value == 'OK':    #TEST CASE PASSED             
                       Message_counter_result_entry.insert(0,"OK")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is OK',CANID_MSG_CTR)				   
                    elif Message_Counter_Result_Sheet.cell(Message_Counter_CANID_row[counter_row],Message_Counter_CANID_column[k]).value == 'CA': #TEST CASE FAILED
                       Message_counter_result_entry.insert(0,"CA")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is CA',CANID_MSG_CTR)				   
                    else:
                       Message_counter_result_entry.insert(0,"RESULT NOT FOUND")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result for %s is NOT FOUND ',CANID_MSG_CTR)
                    counter_row=counter_row+1
                    progressbar=progressbar+1
                    Message_counter_progressbar['value']=progressbar
                    time.sleep(2)                                 # time delay to keep the value in display boxes fixed
                    

       
                overall_progressbar=overall_progressbar+1
                Message_counter_overall_progressbar["value"]= overall_progressbar
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')
            logging.info('##############  End of Message Counter  Testing  ##############')
        except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')

            logging.exception('Test case execution stopped abrubtly')
            
                       
#**********************************************************************************************************************************************************************#


    def Gateway_TGW_Testing() :

        global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DIAG_CMD_NO_path_3,DTC_string_temp,DTC_string_1_temp,DTC_string_path_temp,Read_vehicle_speed_path        
        global myAppl,Gateway_overall_progressbar,Gateway_progressbar
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,format='%(message)s')
        logging.info('##############  Start of Start of Gateway TGW Testing  ##############')
        
        Gateway_TGW_Dict = OrderedDict()                    # This makes dictionary required for GAteway TGW Tree
        Gateway_TGW_Dict[VehicleName]= OrderedDict()
        Gateway_TGW_Dict[VehicleName][RegionName]= OrderedDict()
        Gateway_TGW_Dict[VehicleName][RegionName][PartNo]= OrderedDict()
        Active_Test = 'GATEWAY_TGW'
        Gateway_TGW_Dict[VehicleName][RegionName][PartNo][Active_Test]=OrderedDict()
        
        Gateway_TGW_Result_Folder=[]     # This contains the path of result folders created individually for all the variants for GateWay result
        Gateway_CANID_list = []      #This contains list of all ENABLED GateWay CANID extracted from the dispatch sheet      
        COUNT_YES = []                       #This stores the number number of Enabled GateWay CANID for each variant. i.e. No of CANID with 'Y' in front of them        
        NO_OF_SET = []
        
        DIMPSheet_Gateway_TGW_CANID_List = WorkBook.sheet_by_index(3)          #This opens the FOURTH Sheet of DISPATCH Sheet i.e.Sheet containing list of a Gateway TGW CANID
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
        logging.info('Dispatch Sheet for Gateway  Loaded')		
        DIMPSheet_Gateway_TGW_CANID_List_Col = DIMPSheet_Gateway_TGW_CANID_List.ncols     
        DIMPSheet_Gateway_TGW_CANID_List_Row = DIMPSheet_Gateway_TGW_CANID_List.nrows
        Gateway_CANID_row=[]     #List used to store row number of "Y" in DISPATCH SHEET WORKBOOK(GATEWAY SHEET).       Used to insert result later in code
        Gateway_CANID_column=[]  #List used to store column number of "Y" in DISPATCH SHEET WORKBOOK(GATEWAY SHEET).    Used to insert result later in code
        Gateway_progressbar["maximum"]=4
        Gateway_overall_progressbar["maximum"]=20
        try:            
            for k in range(0,len(Variant)):
                counter = 0   
                Gateway_TGW_Dict[VehicleName][RegionName][PartNo][Active_Test][Variant[k]]=OrderedDict()
                Gateway_TGW_Result_Folder.append(VehicleNameFolder[k]+"\\08_Gateway\\Gateway_TGW")              # This is Destination folder to store copy of GATEWAY CANape configuration
                dest_Gateway_Result_Folder = Gateway_TGW_Result_Folder[k]
                
                #distutils.dir_util.copy_tree(Master_CANape_Gateway_TGW_Path,dest_Gateway_Result_Folder)            # This function copies files and folders from MASTER CANAPE CONFIG 
                
                copy_folder(Master_CANape_Gateway_TGW_Path,dest_Gateway_Result_Folder)
             
                Var_str =  'Variant_0' +  str(k +1)    #This is used to obtain Variant number in format "Variant_0X" 
                
                Var_Row = 0    #Store row number of Variant in DISPATCH SHEET WORKBOOK(GATEWAY SHEET) 
                Var_col = 0    #Store column number of Variant in DISPATCH SHEET WORKBOOK(GATEWAY SHEET)
                CANID_row=0     #Store row number of the CANID in DISPATCH SHEET WORKBOOK(GATEWAY SHEET)
                CANID_col=0    #Store column number of the CANID in DISPATCH SHEET WORKBOOK(GATEWAY SHEET)
                Node_col = 0
                Node_row = 0

                
                for i in range (0, DIMPSheet_Gateway_TGW_CANID_List_Row):             # Loop for extracting the Vehicle Info
                    for j in range(0, DIMPSheet_Gateway_TGW_CANID_List_Col):
                        if DIMPSheet_Gateway_TGW_CANID_List.cell(i,j).value == Var_str:         #This finds row and column of the particular variant
                            Var_col =  j
                            Var_Row = i
                        if DIMPSheet_Gateway_TGW_CANID_List.cell(i,j).value == 'Node':        #This finds row and column of CANIDs in GATEWAY_DETAIL
                            Node_col = j
                            Node_row = i                            
                        if DIMPSheet_Gateway_TGW_CANID_List.cell(i,j).value == 'CANID':        #This finds row and column of CANIDs in GATEWAY_DETAIL
                             CANID_col =  j
                             CANID_row = i
                        else:
                            continue
                for i in range (Var_Row + 2, DIMPSheet_Gateway_TGW_CANID_List_Row):    # ( Var_Row +2 ) contains the string "Y" or "N"                   
                    if (DIMPSheet_Gateway_TGW_CANID_List.cell(i,Var_col).value=='Y'):
                        CANID_Gateway = DIMPSheet_Gateway_TGW_CANID_List.cell(i,CANID_col).value    #This extracts the CANID form the sheet
                        CANID_Gateway_Node_dir = DIMPSheet_Gateway_TGW_CANID_List.cell(i,Node_col).value    #This extracts the CANID form the sheet
                        if CANID_Gateway in Gateway_TGW_Dict[VehicleName][RegionName][PartNo][Active_Test][Variant[k]]:
                            print 'Already in Dict'
                        else:
                            counter = counter +1           #This counts the number of 'Y' for a particular variant.                
                            Gateway_TGW_Dict[VehicleName][RegionName][PartNo][Active_Test][Variant[k]][CANID_Gateway + "  "+ CANID_Gateway_Node_dir]=OrderedDict()   #This adds CANID to the Gateway Tree                        
                        Gateway_CANID_list.append(DIMPSheet_Gateway_TGW_CANID_List.cell(i,CANID_col).value);    #This adds CANID to a list 
                        Gateway_CANID_row.append(i)   #This stores the row number of all "Y" for future reference
                Gateway_CANID_column.append(Var_col)  #This store the column number of all "Y" for future reference       
                COUNT_YES.append(counter)              #This is the final count of "Y". Each element of COUNT_YES correspons to the corresponding Variant
                print COUNT_YES
                print 'count y'
            print Gateway_TGW_Dict
            
            

            uid_MSG_prev=uid
      
            Gateway_TGW_Tree=construct_JSON_tree(Gateway_TGW_Dict,frame9)
            curItem_Gateway= uid_MSG_prev+1     #Since uid doest become 0
            Gateway_vehicle_id = Gateway_TGW_Tree.item(curItem_Gateway, 'text')
            Gateway_vehicle_id_entry["state"] = NORMAL
            Gateway_vehicle_id_entry.insert(0, Gateway_vehicle_id)
            Gateway_vehicle_id_entry["state"] = DISABLED
            curItem_Gateway=curItem_Gateway+2

            Gateway_variant_id = Gateway_TGW_Tree.item(curItem_Gateway, 'text')
            curItem_Gateway=curItem_Gateway+1        

     
            counter_row=0
            myAppl.Variable(Power_Supply_path).Write(1) 
        
            
            for k in range(0,len(Variant)):
               
                Gateway_variant_entry.delete(0,END)   
                Gateway_CAN_ID_entry.delete(0,END)
                Gateway_result_entry.delete(0,END)
                Gateway_GWTGW_entry.delete(0,END)
                Var_str =  'Variant_0' +  str(k +1)
                
                curItem_Gateway = curItem_Gateway + 1
                Gateway_TGW_Tree.selection_set(curItem_Gateway)
       
                Variant_Name = Gateway_TGW_Tree.item(curItem_Gateway,'text')
                Gateway_variant_entry["state"] = NORMAL
                Gateway_GWTGW_entry["state"] = NORMAL
                Gateway_variant_entry.insert(0,Variant_Name)

                Gateway_GWTGW_entry.insert(0,"GATEWAY_TGW")            
                Gateway_GWTGW_entry["state"] = DISABLED
                Gateway_variant_entry["state"] = DISABLED
                Write_Var = Variant_Value[k]
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('Variant %s under Message Counter testing',Write_Var)			

                myAppl.Variable(Power_Supply_path).Write(129)
                myAppl.Variable(Power_Supply_path).Write(1)   # Switch on Power Supply to Write Variant Code     
                        ##Instrumentation().ActiveLayout.Normalize()


                time.sleep(7)
                
                Variant_write(myAppl,Write_Var) 
                
                myAppl.Variable(Power_Supply_path).Write(129) 
                myAppl.Variable(Power_Supply_path).Write(1) 
 

 
                myAppl = None
                Instrumentation().AnimationMode =0
                time.sleep(2)
                PlatformManager().Platforms.Item(PlatformName).Stop()
                time.sleep(4)

                GatewayIDTextFilepath=Gateway_TGW_Result_Folder[k]+"\\"+"GATEWAYID.txt"      # create gatewayID text file   
                xlapp = win32com.client.Dispatch("Excel.Application")
                if os.path.exists(str(Gateway_Trace_Path)):
                  print "PATH exists"
                  xlapp.Workbooks.Open(Filename=str(Gateway_Trace_Path), ReadOnly=1)
                  xlapp.Application.Run("Gateway_Judgement.xls!module4.xcelltotextconvert",str(GatewayIDTextFilepath),str(Result_DispatchSheet_Path))           # extract  GATEWAYID from dispatch sheet to text file for capl
                  xlapp.Application.Quit() # Comment this out if your excel script closes
                time.sleep(10)

                SyncPathTextFile=Gateway_TGW_Result_Folder[k]+"\\"+"Sync.txt"
                Start_CANape(Gateway_TGW_Result_Folder[k])
                Gateway_progressbar["value"]=1
              
                valueRead_MSG = 1
                while(valueRead_MSG != '8'):                    #Wait for CAPle to tell python that it has finished CANape and other processing work like  MDF conversion

                 syncFileRead = open(SyncPathTextFile,'r')   #CAPle will write '8' in Sync.txt to inform python that it has finished its work
                 valueRead_MSG = syncFileRead.read()
                 syncFileRead.close()
                 time.sleep(1);                
                    #os.remove(Gateway_TGW_Result_Folder[k]+cfg_file)
                    
                new_filename = "\\"+ str(Variant_Name)  + "_GW_TGW" + ".txt"
                TGW_CANape_LOG_file = Gateway_TGW_Result_Folder[k] + new_filename                
                SRC_2= Gateway_TGW_Result_Folder[k] + "\\LOG.txt"
                Gateway_progressbar["value"]=2
                           
                os.rename(str(SRC_2),str(TGW_CANape_LOG_file))
##               
                
                xlapp = win32com.client.Dispatch("Excel.Application")
                if os.path.exists(str(Gateway_Trace_Path)):
                  print "PATH exists"
                  xlapp.Workbooks.Open(Filename=str(Gateway_Trace_Path), ReadOnly=1)
                  xlapp.Application.Run("Gateway_Judgement.xls!module4.TXTconvertXLS",str(TGW_CANape_LOG_file), str(k+1),Gateway_TGW_Result_Folder[k])        # convert log which is in text form to excel
                  xlapp.Application.Quit() # Comment this out if your excel script closes
                time.sleep(10)

                
##                           
                new_filename_trace_TGW = "\\"+ str(Var_str)+"_GW_TGW"+ ".xls"
                Gateway_TGW_Log_Path = Gateway_TGW_Result_Folder[k] + new_filename_trace_TGW                
                              
                xlapp = win32com.client.Dispatch("Excel.Application")
                if os.path.exists(str(Gateway_Trace_Path)):
                    
                    xlapp.Workbooks.Open(Filename=str(Gateway_Trace_Path), ReadOnly=1)
                    xlapp.Application.Run("Gateway_Judgement.xls!module3.oneachsheet",str(Result_DispatchSheet_Path),str(Var_str),"GW_TGW",str(Gateway_TGW_Log_Path),str(Variant_Name))
                    xlapp.Application.Quit() # Comment this out if your excel script closes
                
                Dispatch_Sheet_Result_WorkBook =xlrd.open_workbook(str(Result_DispatchSheet_Path),formatting_info=True)
                Gateway_TGW_Result_Sheet=Dispatch_Sheet_Result_WorkBook.sheet_by_index(3)

                for t in range(0,COUNT_YES[k]):
                    Gateway_result_entry["state"] = NORMAL
                    Gateway_CAN_ID_entry["state"] = NORMAL
                    Gateway_result_entry.delete(0,END)
                    Gateway_CAN_ID_entry.delete(0,END)
                    curItem_Gateway=curItem_Gateway+1
                    Gateway_TGW_Tree.selection_set(curItem_Gateway)
                    CANID_Gateway = Gateway_TGW_Tree.item(curItem_Gateway, 'text')
                    Gateway_CAN_ID_entry.insert(0, CANID_Gateway )
                    Gateway_progressbar["value"]=3
##               
                    #To display Test Case result on Message Counter GUI. Logic is as follows
                    #Message_Counter_CANID_row[counter_row]   Message_Counter_CANID_row is a list obtained earlier containing row number of "Y".
                    #Message_Counter_CANID_column  is a list obtained earlier containing column number of "Y".
                    #"Y" is replaced by "OK"  or "CA" or "RESULT NOT FOUND"    . So row number obtained earlier is utilized
                    Gateway_TGW_Tree.item(curItem_Gateway, text = CANID_Gateway, values = Gateway_TGW_Result_Sheet.cell(Gateway_CANID_row[counter_row],Gateway_CANID_column[k]).value)  # Displaying result in the GUI
                    if Gateway_TGW_Result_Sheet.cell(Gateway_CANID_row[counter_row],Gateway_CANID_column[k]).value == 'OK':     #TEST CASE PASSED            
                       Gateway_result_entry.insert(0,"OK")                       
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is OK',CANID_Gateway)

                  
                       
                    elif Gateway_TGW_Result_Sheet.cell(Gateway_CANID_row[counter_row],Gateway_CANID_column[k]).value == 'CA':   #TEST CASE FAILED 
                       Gateway_result_entry.insert(0,"CA")
                       
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is CA',CANID_Gateway)
                       
                    else:
                       Gateway_result_entry.insert(0,"RESULT NOT FOUND")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is NOT FOUND',CANID_Gateway)                   
                       
                    counter_row=counter_row+1
                    time.sleep(2)                 # time delay to keep the value in display boxes fixed

                    Gateway_progressbar["value"]=4
                dest_files = os.listdir(Gateway_TGW_Result_Folder[k])
                for file in dest_files:
                    if file[-4:] == ".exe" or file[-4:] == ".ctf" or \
                       file[-4:] == ".BLF" or file[-4:] == ".scr" or file[-2:] == ".c" or \
                       file == "GATEWAYID.txt"  or \
                       file == "Sync.txt" :
                        os.remove(Gateway_TGW_Result_Folder[k] + "\\" + file)                    

                myAppl = rtplib.Appl(ApplFile + ".sdf", PlatformName, SystemType)
                time.sleep(4)
                PlatformManager().Platforms.Item(PlatformName).Start()
                Instrumentation().AnimationMode =2
                Gateway_overall_progressbar["value"]= ((20*(k+1))/(len(Variant)))
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,format='%(message)s')
            logging.info('##############  End of Start of Gateway TGW Testing  ##############')
        except Exception, e:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')

            logging.exception('Test case execution stopped abrubtly')
###################################################################################################################################################################                          
            

######################################################################Busoff Check Testing #######################################################################################################################

    def BusOff_Testing() :
        global myAppl

        global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DIAG_CMD_NO_path_3,DTC_string_temp,DTC_string_1_temp,DTC_string_path_temp,Read_vehicle_speed_path
        time.sleep(3)

        overall_progressbar=0
        Vehicle_Details = VehicleName +  '_' + RegionName
        Vehicle_Name = VehicleName + '_' + RegionName + '_' + PartNo
        Actual_PartNo=PartNo.split("_")[1]
        BusOff_overall_progressbar["value"] = overall_progressbar

        logging.info('##############  Start of Busoff Check  Testing  ##############')
        BusOff_Dict = OrderedDict()                    # This makes dictionary required for Busoff Tree
        BusOff_Dict[Vehicle_Details]= OrderedDict()

        BusOff_Dict[Vehicle_Details][PartNo]= OrderedDict()
        Active_Test = 'Busoff_Check'

        BusOff_Dict[Vehicle_Details][PartNo][Active_Test]=OrderedDict()

        CAN_Channel_Array_list=[]
        BusOff_Result_Folder=[]     # This contains the path of result folders created individually for all the variants for Busoff result
        BusOff_CANID_list = []      #This contains list of all ENABLED Busoff CANID extracted from the dispatch sheet
        COUNT_YES = []                       #This stores the number number of Enabled Busoff CANID for each variant. i.e. No of CANID with 'Y' in front of them



        DIMPSheet_BusOff_CANID_List = WorkBook.sheet_by_index(5)          #This opens the SECOND Sheet of DISPATCH Sheet i.e.Sheet containing list of active Message Counter CANID
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
        logging.info('Dispatch Sheet for Busoff Check Loaded')
        DIMPSheet_BusOff_CANID_List_Col = DIMPSheet_BusOff_CANID_List.ncols
        DIMPSheet_BusOff_CANID_List_Row = DIMPSheet_BusOff_CANID_List.nrows
        BusOff_CANID_row=[]     #List used to store row number of "Y" in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET).       Used to insert result later in code
        BusOff_CANID_column=[]  #List used to store column number of "Y" in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET).    Used to insert result later in code



        BusOff_overall_progressbar["value"]=overall_progressbar
        for k in range(0,len(Variant)):
            CAN_Channel_Array_list=[]
            counter = 0
            data  = Variant[k].split('_')
            Variant_Number_string = data[2]
            VariantName_tree = ' Variant ' + data[2]
            BusOff_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree]=OrderedDict()
            BusOff_Result_Folder.append(VehicleNameFolder[k]+"\\04_Bus_Off_Check")              # This is Destination folder to store copy of BusOff CANape configuration
            dest_BusOff_Result_Folder = BusOff_Result_Folder[k]

            distutils.dir_util.copy_tree(Master_CANape_BusOff_Path,dest_BusOff_Result_Folder)
            ##copy_folder(Master_CANape_BusOff_Path,dest_BusOff_Result_Folder)        # This function copies files and folders from MASTER CANAPE CONFIG

            Var_str =  'Variant_0' +  str(k +1)    #This is used to obtain Variant number in format "Variant_0X" eg. Variant_01 Variant_02 etc

            Var_Row = 0     #Store row number of Variant in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)
            Var_col = 0     #Store column number of Variant in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)
            CANID_row=0     #Store row number of the CANID in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)
            CANID_col=0     #Store column number of the CANID in DISPATCH SHEET WORKBOOK(MESSAGE COUNTER SHEET)


            for i in range (0,  DIMPSheet_BusOff_CANID_List_Row):             # Loop for traversing through the EXCEL sheet
                for j in range(0,DIMPSheet_BusOff_CANID_List_Col):
                    if DIMPSheet_BusOff_CANID_List.cell(i,j).value == Var_str:         #This finds row and column of the particular variant
                        Var_col =  j
                        Var_Row = i

                    if DIMPSheet_BusOff_CANID_List.cell(i,j).value == 'CANID':        #This finds row and column of CANIDs in MSG_COUNTER_DETAIL
                         CANID_col =  j
                         CANID_row = i
                    if DIMPSheet_BusOff_CANID_List.cell(i,j).value == 'CAN Channel':        #This finds row and column of CANIDs in MSG_COUNTER_DETAIL
                         CAN_Channel_col =  j
                         CAN_Channel_row = i
                    else:
                        continue



            for i in range (Var_Row + 2, DIMPSheet_BusOff_CANID_List_Row):      # ( Var_Row +2 ) contains the string "Y" or "N"

                if (DIMPSheet_BusOff_CANID_List.cell(i,Var_col).value=='Y'):

                    counter = counter +1           #This counts the number of 'Y' for a particular variant.
                    CAN_Channel_value=DIMPSheet_BusOff_CANID_List.cell(i,CAN_Channel_col).value
                    CANID_BusOff=DIMPSheet_BusOff_CANID_List.cell(i,CANID_col).value    #This extracts the CANID form the sheet

                    BusOff_CANID_list.append(DIMPSheet_BusOff_CANID_List.cell(i,CANID_col).value)   #This adds CANID to a list
                    BusOff_CANID_row.append(i)   #This stores the row number of all "Y" for future reference
                    if CAN_Channel_value in CAN_Channel_Array_list:
                        BusOff_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][CAN_Channel_value][CANID_BusOff] = OrderedDict()

                    else :
                        CAN_Channel_Array_list.append(CAN_Channel_value)
                        BusOff_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][CAN_Channel_value] = OrderedDict()
                        BusOff_Dict[Vehicle_Details][PartNo][Active_Test][VariantName_tree][CAN_Channel_value][CANID_BusOff] = OrderedDict()

            BusOff_CANID_column.append(Var_col)  #This store the column number of all "Y" for future reference


            COUNT_YES.append(counter)                      #This is the final count of "Y". Each element of COUNT_YES correspons to the corresponding Variant



        uid_MSG_prev=uid
        BusOff_Tree = construct_JSON_tree(BusOff_Dict,frame10)    #This makes Message Counter Tree
        curItem_MSG_CTR= uid_MSG_prev+1     #Since uid doest become 0
        BusOff_vehicle_id = Vehicle_Name
        Busoff_vehicle_id_entry.insert(0, BusOff_vehicle_id)            #Fills the space in Message Counter GUI with required information

        curItem_MSG_CTR=curItem_MSG_CTR+2   #This is used to jumpover the items not required in Tree for displaying in GUI

        counter_row=0 #Simple counter variable for incrementing



        for k in range(0,len(Variant)):
            data1  = Variant[k].split('_')
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
            logging.info('data1 %s ',data1)
            Variant_Number_string = data1[2]
            logging.info('Variant_Number_string %s ',Variant_Number_string)
            logging.info('k %s ',k)
            progressbar=0
            BusOff_progressbar["maximum"] = (COUNT_YES[k])+7
            BusOff_overall_progressbar["maximum"] = len(Variant)
            BusOff_overall_progressbar["value"] = overall_progressbar
            BusOff_progressbar['value']=progressbar

            InterfacevalueRead_MSG=1

            Busoff_variant_entry.delete(0,END)      #Clears the space in Message Counter GUI
            Busoff_CAN_ID_entry.delete(0,END)       #Clears the space in Message Counter GUI
            Busoff_result_entry.delete(0,END)       #Clears the space in Message Counter GUI

            SyncPathTextFile = BusOff_Result_Folder[k] + "\\" + "Sync.txt" #Sync.txt is used to sync PYTHON and CANAPE (CAPL script)
            CSVFile = BusOff_Result_Folder[k] + "\\" + "CANape.txt"        #CANape.txt stores data converted from MDF file. Used for judgement


            Var_str =  'Variant_0' +  str(k +1)  #This is used to obtain Variant number in format "Variant_0X" eg. Variant_01 Variant_02 etc

            curItem_MSG_CTR = curItem_MSG_CTR + 1
            BusOff_Tree.selection_set(curItem_MSG_CTR)   #Highlighting the item in the Tree

            Variant_Name = BusOff_Tree.item(curItem_MSG_CTR, 'text')
            Busoff_variant_entry.insert(0, Variant_Name)

            Write_Var = Variant_Value[k]

            myAppl.Variable(Power_Supply_path).Write(1)   # Switch on Power Supply to Write Variant Code
                    ##Instrumentation().ActiveLayout.Normalize()
            Variant_write(myAppl,Write_Var)
            progressbar=progressbar+1
            BusOff_progressbar['value']=progressbar             
            myAppl.Variable(Power_Supply_path).Write(1)
            BusOff_Result_Folder_pass=BusOff_Result_Folder[k]

            print BusOff_Result_Folder_pass
            Var_str =  'Variant_0' +  str(k +1)    #This is used to obtain Variant number in format "Variant_0X" eg. Variant_01 Variant_02 etc

            Busoff_Check_Report = Org_Path + '05_Master_Result_Reports' + "\\" + "Busoff_Check" + "\\" + 'Busoff_Check_Report.xls'
            Renamed_Busoff_Check_Report=  Actual_PartNo + "_"+ Variant_Number_string + "_" + "Busoff_Report.xls"

            Busoff_Check_Report_path=BusOff_Result_Folder_pass+"\\"+Renamed_Busoff_Check_Report

            shutil.copy(Busoff_Check_Report,Busoff_Check_Report_path)

            CAR_SLCT_NO_read = myAppl.Variable('Model Root/Environment/CarSetting/CAR_SLCT_NO/Value').Read()
            CAR_SLCT_NO_read=str(CAR_SLCT_NO_read)
            if(CAN1_Busoff_Enabled==1):

                BusOff_CAN1_Testing(BusOff_Result_Folder_pass,Write_Var,Var_str,Variant_Number_string,Busoff_Check_Report_path,CAR_SLCT_NO_read,k)
                progressbar=progressbar+3
                BusOff_progressbar['value']=progressbar
            if(CAN2_Busoff_Enabled==1):

                BusOff_CAN2_Testing(BusOff_Result_Folder_pass,Write_Var,Var_str,Variant_Number_string,Busoff_Check_Report_path,CAR_SLCT_NO_read,k)
                progressbar=progressbar+3
                BusOff_progressbar['value']=progressbar
                
            Dispatch_Sheet_Result_WorkBook =xlrd.open_workbook(str(Result_DispatchSheet_Path),formatting_info=True) #This opens Result DISPATCH SHEET Workbook
            BusOff_Result_Sheet=Dispatch_Sheet_Result_WorkBook.sheet_by_index(5)   #The Message Counter Sheet in DISPATCH SHEET WORKBOOK
            

            for t in range(0,(COUNT_YES[k]+len(Number_of_BusOff_CAN_Channel))):
                Busoff_result_entry.delete(0,END)   #Clears the space in Message Counter GUI
                Busoff_CAN_ID_entry.delete(0,END)   #Clears the space in Message Counter GUI
                curItem_MSG_CTR=curItem_MSG_CTR+1            #

                CANID_MSG_CTR = BusOff_Tree.item(curItem_MSG_CTR, 'text')
                if (CANID_MSG_CTR == "CAN 1"):
                    Busoff_CANchannel_entry.delete(0,END)
                    BusOff_Tree.selection_set(curItem_MSG_CTR)
                    Busoff_CANchannel_entry.insert(0, CANID_MSG_CTR )
                    time.sleep(2)



                elif (CANID_MSG_CTR == "CAN 2"):
                    Busoff_CANchannel_entry.delete(0,END)
                    BusOff_Tree.selection_set(curItem_MSG_CTR)
                    Busoff_CANchannel_entry.insert(0, CANID_MSG_CTR )
                    time.sleep(2)


                else:


                    BusOff_Tree.selection_set(curItem_MSG_CTR)
                    Busoff_CAN_ID_entry.insert(0, CANID_MSG_CTR )
                    #To display Test Case result on Message Counter GUI. Logic is as follows
                    #BusOff_CANID_row[counter_row]   BusOff_CANID_row is a list obtained earlier containing row number of "Y".
                    #BusOff_CANID_column  is a list obtained earlier containing column number of "Y".
                    #"Y" is replaced by "OK"  or "CA" or "RESULT NOT FOUND"    . So row number obtained earlier is utilized
                    if BusOff_Result_Sheet.cell(BusOff_CANID_row[counter_row],BusOff_CANID_column[k]).value == 'OK':    #TEST CASE PASSED
                       Busoff_result_entry.insert(0,"OK")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is OK',CANID_MSG_CTR)
                    elif BusOff_Result_Sheet.cell(BusOff_CANID_row[counter_row],BusOff_CANID_column[k]).value == 'CA': #TEST CASE FAILED
                       Busoff_result_entry.insert(0,"CA")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result of %s is CA',CANID_MSG_CTR)
                    else:
                       Busoff_result_entry.insert(0,"RESULT NOT FOUND")
                       logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                       logging.info('Result for %s is NOT FOUND ',CANID_MSG_CTR)
                    counter_row=counter_row+1
                    progressbar=progressbar+1
                    BusOff_progressbar['value']=progressbar
                    time.sleep(2)                                 # time delay to keep the value in display boxes fixed



            overall_progressbar=overall_progressbar+1
            BusOff_overall_progressbar["value"]= overall_progressbar
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(message)s')
        logging.info('##############  End of Busoff Check  Testing  ##############')






    def BusOff_CAN1_Testing(BusOff_Result_Folder_pass,Write_Var,Var_str,Variant_Number_string,Busoff_Check_Report_path,CAR_SLCT_NO_read,k):
        global myAppl
        BusOff_Result_Folder_CAN1=BusOff_Result_Folder_pass+ "\\"+ "CAN1_Busoff"
        SyncPathTextFile=BusOff_Result_Folder_CAN1+"\\"+"Sync.txt"
        Start_CANape(BusOff_Result_Folder_CAN1)
        Variant_No = k
        valueRead_MSG = 1
        while(valueRead_MSG != '7'):                    #Wait for CAPle to tell python that it has finished CANape and other processing work like  MDF conversion

            syncFileRead = open(SyncPathTextFile,'r')   #CAPle will write '8' in Sync.txt to inform python that it has finished its work
            valueRead_MSG = syncFileRead.read()
            syncFileRead.close()
            time.sleep(1);

        time.sleep(10)
        myAppl.Variable(Power_Supply_path).Write(14)
        time.sleep(15)

        myAppl = None
        Instrumentation().AnimationMode =0
        time.sleep(2)
        PlatformManager().Platforms.Item(PlatformName).Stop()
        time.sleep(4)
        myAppl = rtplib.Appl(ApplFile + ".sdf", PlatformName, SystemType)
        time.sleep(4)
        PlatformManager().Platforms.Item(PlatformName).Start()
        Instrumentation().AnimationMode =2

        time.sleep(2)
        myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)
        time.sleep(1)

        time.sleep(5)

        sync_num = 8                                      #Write '9' in Sync.txt to tell CAPle to start working
        syncFileWriteMSG = open(SyncPathTextFile,'w')
        valueWrite= str(sync_num)
        syncFileWriteMSG.write(valueWrite)
        syncFileWriteMSG.close()

        time.sleep(2)

        xlapp = win32com.client.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet

        if os.path.exists(str(BUSOFF_JUDGEMENT_SHEET)):

            xlapp.Workbooks.Open(Filename=str(BUSOFF_JUDGEMENT_SHEET), ReadOnly=1)
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')
            logging.info('Judgement for BusOff Check [CAN1] started')

            print "Result_DispatchSheet_Path" , Result_DispatchSheet_Path
            ##print "CSVFile" , CSVFile
            ##print "InterfaceTextFile" , InterfaceTextFile
            xlapp.Application.Run("Busoff_Judgement_Sheet.xls!module4.Main_Bus_off", 'CAN1_Busoff_Log',str(BusOff_Result_Folder_CAN1),str(Result_DispatchSheet_Path),Var_str)
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')
            logging.info('Judgement for Busoff Check [CAN1] finished')

            xlapp.Application.Quit() # Comment this out if your excel script closes

        nb= BusOff_Result_Folder_CAN1 +"\\Format.txt"
        with open(nb, 'r+') as f:
            first_line = f.readline()
        first_line = first_line.replace('"','')
        f = open(nb, 'w+')
        f.write(first_line)
        f.truncate()
        f.close()


        write_ss_info = open(BusOff_Result_Folder_CAN1 + "\\Screen_Shot.txt","w")
        write_ss_info.write(BusOff_Result_Folder_pass + "\\" + "00_Screenshots\n")
        write_ss_info.write("01")
        write_ss_info.write("_")
        write_ss_info.write("Busoff_V_CAN")
        write_ss_info.close()

        sync_num = 9                                      #Write '9' in Sync.txt to tell CAPle to start working
        syncFileWriteMSG = open(SyncPathTextFile,'w')
        valueWrite= str(sync_num)
        syncFileWriteMSG.write(valueWrite)
        syncFileWriteMSG.close()

        while(valueRead_MSG != '10'):                    #Wait for CAPle to tell python that it has finished CANape and other processing work like  MDF conversion

            syncFileRead = open(SyncPathTextFile,'r')   #CAPle will write '8' in Sync.txt to inform python that it has finished its work
            valueRead_MSG = syncFileRead.read()
            syncFileRead.close()
            time.sleep(1);
        Busoff_Check_Screenshot_Path=BusOff_Result_Folder_pass + "\\" + "00_Screenshots"

        print Var_str
        print Busoff_Check_Screenshot_Path
        print CAR_SLCT_NO_read
        print VehicleName
        print Variant_Number_string
        print Result_DispatchSheet_Path
        print Busoff_Check_Report_path
        print RegionName
        temp = str(Busoff_Check_Screenshot_Path)
        temp1= str(Variant_Number_string)

        xlapp = win32com.client.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet

        if os.path.exists(str(BUSOFF_JUDGEMENT_SHEET)):

            xlapp.Workbooks.Open(Filename=str(BUSOFF_JUDGEMENT_SHEET), ReadOnly=1)
            xlapp.Application.Run("Busoff_Judgement_Sheet.xls!module5.report_Generation", 'V_CAN',Var_str,str(Busoff_Check_Screenshot_Path),CAR_SLCT_NO_read,VehicleName,str(Variant_Number_string),str(Result_DispatchSheet_Path),Busoff_Check_Report_path,RegionName,Variant_No)


            xlapp.Application.Quit() # Comment this out if your excel script closes


        myAppl.Variable(Power_Supply_path).Write(129)
        time.sleep(5)
        os.remove(BusOff_Result_Folder_CAN1 + "\\CANape_Script_V3.scr")
        os.remove(BusOff_Result_Folder_CAN1 + "\\PrintScreen.exe")
        #dest_files = os.listdir(BusOff_Result_Folder_CAN1)
        #for file in dest_files:
         #   if file[-4:] == ".exe" or \
          #     file[-4:] == ".scr" or \
           #    file == "Screen_Shot.txt" or \
            #   file == "sync.txt" :
             #   os.remove(dest_folder + "\\" + file)

    def BusOff_CAN2_Testing(BusOff_Result_Folder_pass,Write_Var,Var_str,Variant_Number_string,Busoff_Check_Report_path,CAR_SLCT_NO_read,k):

        global myAppl
        BusOff_Result_Folder_CAN2=BusOff_Result_Folder_pass+ "\\"+ "CAN2_Busoff"
        SyncPathTextFile=BusOff_Result_Folder_CAN2+"\\"+"Sync.txt"
        myAppl.Variable(Power_Supply_path).Write(1)
        Start_CANape(BusOff_Result_Folder_CAN2)
        Variant_No = k

        valueRead_MSG = 1
        while(valueRead_MSG != '7'):                    #Wait for CAPle to tell python that it has finished CANape and other processing work like  MDF conversion

            syncFileRead = open(SyncPathTextFile,'r')   #CAPle will write '8' in Sync.txt to inform python that it has finished its work
            valueRead_MSG = syncFileRead.read()
            syncFileRead.close()
            time.sleep(1);

        time.sleep(10)
        myAppl.Variable(Power_Supply_path).Write(15)
        time.sleep(10)

        myAppl = None
        Instrumentation().AnimationMode =0
        time.sleep(2)
        PlatformManager().Platforms.Item(PlatformName).Stop()
        time.sleep(4)
        myAppl = rtplib.Appl(ApplFile + ".sdf", PlatformName, SystemType)
        time.sleep(4)
        PlatformManager().Platforms.Item(PlatformName).Start()
        Instrumentation().AnimationMode =2

        time.sleep(2)
        myAppl.Variable(CAR_SLCT_NO_path).Write(Write_Var)
        time.sleep(6)

        sync_num = 8                                      #Write '9' in Sync.txt to tell CAPle to start working
        syncFileWriteMSG = open(SyncPathTextFile,'w')
        valueWrite= str(sync_num)
        syncFileWriteMSG.write(valueWrite)
        syncFileWriteMSG.close()

        time.sleep(2)

        xlapp = win32com.client.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet


        if os.path.exists(str(BUSOFF_JUDGEMENT_SHEET)):

            xlapp.Workbooks.Open(Filename=str(BUSOFF_JUDGEMENT_SHEET), ReadOnly=1)
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')
            logging.info('Judgement for BusOff Check [CAN2] started')

            print "Result_DispatchSheet_Path" , Result_DispatchSheet_Path
            ##print "CSVFile" , CSVFile
            ##print "InterfaceTextFile" , InterfaceTextFile
            xlapp.Application.Run("Busoff_Judgement_Sheet.xls!module4.Main_Bus_off", 'CAN2_Busoff_Log',str(BusOff_Result_Folder_CAN2),str(Result_DispatchSheet_Path),Var_str)
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')
            logging.info('Judgement for Busoff Check [CAN2] finished')

            xlapp.Application.Quit() # Comment this out if your excel script closes

        nb= BusOff_Result_Folder_CAN2 +"\\Format.txt"
        with open(nb, 'r+') as f:
            first_line = f.readline()
        first_line = first_line.replace('"','')
        f = open(nb, 'w+')
        f.write(first_line)
        f.truncate()
        f.close()


        write_ss_info = open(BusOff_Result_Folder_CAN2 + "\\Screen_Shot.txt","w")
        write_ss_info.write(BusOff_Result_Folder_pass + "\\" + "00_Screenshots\n")
        write_ss_info.write("02")
        write_ss_info.write("_")
        write_ss_info.write("Busoff_ITS_CAN")
        write_ss_info.close()

        sync_num = 9                                      #Write '9' in Sync.txt to tell CAPle to start working
        syncFileWriteMSG = open(SyncPathTextFile,'w')
        valueWrite= str(sync_num)
        syncFileWriteMSG.write(valueWrite)
        syncFileWriteMSG.close()
        Busoff_Check_Screenshot_Path=BusOff_Result_Folder_pass + "\\" + "00_Screenshots"

        while(valueRead_MSG != '10'):                    #Wait for CAPle to tell python that it has finished CANape and other processing work like  MDF conversion

            syncFileRead = open(SyncPathTextFile,'r')   #CAPle will write '8' in Sync.txt to inform python that it has finished its work
            valueRead_MSG = syncFileRead.read()
            syncFileRead.close()
            time.sleep(1);


        xlapp = win32com.client.Dispatch("Excel.Application")   #To open Excel for Message Counter Judgement Sheet

        if os.path.exists(str(BUSOFF_JUDGEMENT_SHEET)):
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')

            logging.info('You have selected the following CAN2 reportif loop before')

            xlapp.Workbooks.Open(Filename=str(BUSOFF_JUDGEMENT_SHEET), ReadOnly=1)
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')

            logging.info('You have selected the following CAN2 reportif loop open')

            xlapp.Application.Run("Busoff_Judgement_Sheet.xls!module5.report_Generation", 'ITS_CAN',Var_str,str(Busoff_Check_Screenshot_Path),CAR_SLCT_NO_read,VehicleName,str(Variant_Number_string),str(Result_DispatchSheet_Path),Busoff_Check_Report_path,RegionName,Variant_No)



            xlapp.Application.Quit() # Comment this out if your excel script closes



        myAppl.Variable(Power_Supply_path).Write(129)
        os.remove(BusOff_Result_Folder_CAN2 + "\\CANape_Script_V3.scr")
        os.remove(BusOff_Result_Folder_CAN2 + "\\PrintScreen.exe")
        
###################################################################################################################################################################

######################################################################Config Check Testing #######################################################################################################################
    def Config_Check_Testing():
        
        global myAppl,Dispatch_dest_Folder,DIAG_CMD_NO_path,Config_check_col_new,Variant_Value,Config_check_enabled_tree_col,dest_config_check_Result_Folder,Test_Sheet_Path
        global Config_check_enabled_tree,Config_check_vehicle_id,ConfigCheck_vehicle_id_entry,ConfigCheck_variant_entry,ConfigCheck_result_entry       

        global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DIAG_CMD_NO_path_3,DTC_string_temp,DTC_string_1_temp,DTC_string_path_temp,Read_vehicle_speed_path
        global Book_Master_TP
        
        uid_config_prev=uid
        curItem_config= uid_config_prev+1    
        read_part_chr=[]                                                                                                            #Array to read part number 
        read_config_chr=[]                                                                                                          #Array to read config number 
        Config_check_enabled_tree_col=[]
        
        Config_Check_Dict = OrderedDict()
        Config_Check_Dict[VehicleName]= OrderedDict()
        Config_Check_Dict[VehicleName][RegionName]= OrderedDict()                                                                   #Config_check dictionary 
        Config_Check_Dict[VehicleName][RegionName][PartNo]= OrderedDict()
       
        logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                level=logging.INFO,
                                format='%(message)s')
        logging.info('##############  Config Check Started  ##############')
        
##        interface_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Interface_VBA.xls'
        excel= win32com.client.dynamic.Dispatch("Excel.Application")                                                                
        interface_vba = excel.Workbooks.Open(str(interface_sheet_path))
        excel.Visible = False
        interface_vba_sheet = interface_vba.Sheets(1)
        interface_vba_sheet.Cells(1,2).Value = VehicleName
        interface_vba_sheet.Cells(2,2).Value = RegionName
        interface_vba.Save()
        interface_vba.Close()
        #Test_Sheet = Test_Sheet_Path + '\\' + VehicleName + '\\' + 'Master_TestPattern_ITS.xls'
                
        try:    
            for k in range(0,len(Variant)):                                                                                             #This loop is for making dictionary for config 
                
                Config_Check_Dict[VehicleName][RegionName][PartNo][Variant[k]]=OrderedDict()        
                Config_Check_Dict[VehicleName][RegionName][PartNo][Variant[k]][Config_check_enabled_tree[k]]=OrderedDict()
            
            Config_check_Tree = construct_JSON_tree(Config_Check_Dict,frame11)  
            Config_check_Result_Folder= []
            time.sleep(1)
            
            Config_check_Tree.selection_set(curItem_config)                                                 #Highlight vehicle name 
            
            Config_check_vehicle_id = Config_check_Tree.item(curItem_config, 'text')
            ConfigCheck_vehicle_id_entry["state"] = NORMAL
            ConfigCheck_vehicle_id_entry.insert(0, Config_check_vehicle_id)
            ConfigCheck_vehicle_id_entry["state"] = DISABLED

            curItem_config=curItem_config+1
            time.sleep(1)
            Config_check_Tree.selection_set(curItem_config)                                                 #Highlight region name 
            curItem_config=curItem_config+1
            time.sleep(1)
            Config_check_Tree.selection_set(curItem_config)                                                 #Highlight Part number
            curItem_config=curItem_config+1
            time.sleep(1)
            overall_progress=1
            overall_max=len(Variant)+2
            ConfigCheck_progressbar["maximum"] = 4
            ConfigCheck_overall_progressbar["maximum"] = overall_max
            ConfigCheck_overall_progressbar["value"] = overall_progress
            print "Config_check_enabled_tree: " + str(Config_check_enabled_tree)
            for k in range(0,len(Variant)):
                
                ConfigCheck_variant_entry["state"] = NORMAL

                ConfigCheck_variant_entry.delete(0,END)                                                                 # Clears the space in Config_check Variant entry GUI  
                
                Config_check_Result_Folder.append(VehicleNameFolder[k]+"\\03_Others_DTC_Check_Config_Check")            # This is Destination folder to store copy of config_check report
                
                print Config_check_report_folder_path
                dest_config_check_Result_Folder = Config_check_Result_Folder[k]
                print dest_config_check_Result_Folder
                time.sleep(0.5)
                            
                ConfigCheck_overall_progressbar["value"] = overall_progress
                overall_progress=overall_progress+1
                time.sleep(0.5)
                WorkBook_dis =xlrd.open_workbook(Result_DispatchSheet_Path,                                     # Open Result Dispatch Sheet
                                                 formatting_info=True)
                    
                DIMSheet=WorkBook_dis.sheet_by_index(0)                                                         # To store first sheet if dispatch sheet in var
                                                                                                                
                DIMSheetCol=DIMSheet.ncols
                DIMSheetRow=DIMSheet.nrows
                for i in range(0,DIMSheetRow):                                                                  # Loop to extract "Config_start"string from result dispatch sheet 
                        for j in range(0,DIMSheetCol):
                            if (DIMSheet.cell(i,j).value=="Config_start"): 
                                config_list_col=j+1
                                DispatchRow=i+k
                
                if (Config_check_enabled_tree[k]!=''):
 
                    distutils.dir_util.copy_tree(Config_check_report_folder_path,dest_config_check_Result_Folder)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.exception('Enable application is found')                                                                               # If "o" is found in dispatch sheet then this condition is true 
                    myAppl.Variable(Power_Supply_path).Write(1)                                                         # Switch on Power Supply to Write Variant Code     
                    Var_str =  'Variant_' +  str(int(Variant_Value[k]))                                                 # Var_str will store the variant number 
                
                    ConfigCheck_variant =Var_str
                    Config_check_Tree.selection_set(curItem_config)
                    time.sleep(0.5)
                    ConfigCheck_result_entry.delete(0,END)
                    ConfigCheck_variant_entry["state"] = NORMAL
                    ConfigCheck_variant_entry.insert(0, ConfigCheck_variant)
                    ConfigCheck_variant_entry["state"] = DISABLED
                    Write_Var = Variant_Value[k]
                    time.sleep(0.5)
                    
                    Variant_write(myAppl,Write_Var)                                                                         #Function to write variant 
                    time.sleep(0.5) 
                    ConfigCheck_progressbar["value"]=1
                    time.sleep(1)               
                    curItem_config=curItem_config+1
                    time.sleep(0.7)
                    Config_check_Tree.selection_set(curItem_config)                                                           #To highlight config_tree                                                    
                    curItem_config=curItem_config+1

                    myAppl.Variable(Power_Supply_path).Write(1)
                    time.sleep(0.5)     
                    myAppl.Variable(DIAG_CMD_NO_path).Write(12)
                    time.sleep(0.6)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(0.6)
                    
                    ConfigCheck_progressbar["value"] = 1.5                
                    
                    if "ID7C3_TX" in DIAG_CMD_NO_path:
                        DiagID = "ID7C9_RX"
                    else:
                        DiagID = "ID77D_RX"
                        
                    Variant_number = myAppl.Variable('Model Root/Environment/CarSetting/FUNC_CAR_INFO/VARIANT_CD').Read()     #Reading Variant number
                    
                    for x in range(1,11):                                                                                   #This loop is used to read Part number
                    
                        part_number = myAppl.Variable("Model Root/Driver Block/CANdb set/DIAG/" + DiagID + "/PRC_READ_ECM_PART_NO/ECU_PART_NO{SubArray" + str(x) + "}").Read()
                        part_number_int=int(part_number)
                        part_number_chr=chr(part_number_int)
                        read_part_chr.append(part_number_chr)

                    read_part_str = ''.join(map (str, read_part_chr))
                    del read_part_chr[:]
                                    
                    time.sleep(0.5)
                    ConfigCheck_progressbar["value"] = 2.5
                    myAppl.Variable(DIAG_CMD_NO_path).Write(13)
                    time.sleep(0.5)
                    myAppl.Variable(DIAG_CMD_NO_path).Write(0)
                    time.sleep(0.5)
                                
                    
                    
                    for x in range(1,11):                                                                                   #This loop is used to read config number 
                        
                        config_number = myAppl.Variable("Model Root/Driver Block/CANdb set/DIAG/" + DiagID + "/PRC_READ_CONFIG_REF/CONFIG_REF{SubArray" + str(x) + "}").Read()
                        config_number_int=int(config_number)
                        config_number_chr=chr(config_number_int)
                        read_config_chr.append(config_number_chr)

                    read_config_str = ''.join(map (str, read_config_chr))   
                    del read_config_chr[:]
                    time.sleep(0.5)
                    ConfigCheck_progressbar["value"] = 3
                    Config_result_workbook_path = dest_config_check_Result_Folder +"\\"+"Config_Check_Report.xls"

                    
                    Report_new_name=dest_config_check_Result_Folder + "\\"+'Config_Check_Report'+'_'+str(int(Variant_Value[k]))+".xls"
                    os.rename(str(Config_result_workbook_path),str(Report_new_name))
                    print "Result_DispatchSheet_Path",Result_DispatchSheet_Path
                                   
                    config_str=DIMSheet.cell(DispatchRow,config_list_col).value                                               #Stores config_number 
            
                    print "config_str_____________",config_str
                    #DispatchRow=DispatchRow+1
                    new_PartNo = PartNo.replace("_", "")
                    
                    excel= win32com.client.dynamic.Dispatch("Excel.Application")                                                #Opens Config_check Report in result folder
                    Config_check_workbook = excel.Workbooks.Open(str(Report_new_name))
                    excel.Visible = False
                    config_check_sheet = Config_check_workbook .Sheets(1)

                    config_check_sheet.Cells(5,2).Value = VehicleName                
                    config_check_sheet.Cells(6,3).Value = k+1
                    config_check_sheet.Cells(6,4).Value = new_PartNo 
                    config_check_sheet.Cells(6,5).Value = config_str
                    config_check_sheet.Cells(6,6).Value = Variant_number
                    config_check_sheet.Cells(6,7).Value = read_part_str
                    config_check_sheet.Cells(6,8).Value = read_config_str
                    config_Result=config_check_sheet.Cells(6,9).Value                                                      #It will write "OK" or "CA"
                    
                    time.sleep(0.5)
                    ConfigCheck_progressbar["value"] = 3.5

                    style_string1 = "align:horizontal center; pattern: pattern solid, fore_colour red;borders: bottom thin,left thin,right thin,top thin;"
                    style1 = xlwt.easyxf(style_string1)

                    style_string2 = "align:horizontal center; pattern: pattern solid, fore_colour green;borders: bottom thin,left thin,right thin,top thin;"
                    style2 = xlwt.easyxf(style_string2)
                    
                    DisCopy_workbook = copy(WorkBook_dis) 
                    FirstSheet = DisCopy_workbook.get_sheet(0)
                    if (config_Result=='CA'):                    
                        FirstSheet.write(Config_check_row,col_arr[k],'o',style1)
                        
                    else:
                        FirstSheet.write(Config_check_row,col_arr[k],'o',style2)
                        
                    DisCopy_workbook.save(Result_DispatchSheet_Path)                                                    #Make copy of result dispatch sheet 
                    
                    Config_check_workbook.Save()
                    Config_check_workbook.Close()
                    
                    ConfigCheck_progressbar["value"]=4
                    time.sleep(0.5)
                    ConfigCheck_result_entry["state"]= NORMAL
                    
                    
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                    
                    logging.info('Config_check Completed for - %s',Var_str)
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                format='%(asctime)s - %(levelname)s - %(message)s')
                    
                    logging.info('Config_check result is - %s',config_Result)
                    
                    ConfigCheck_result_entry.insert(0, config_Result)
                    time.sleep(1)
                    Config_check_Tree.item(curItem_config - 1, text = "Config_check_Testing", values = config_Result )
                    time.sleep(1)
                    
                else:
    ##                shutil.rmtree(VehicleNameFolder[k],onerror=None)
    ##                DispatchRow=DispatchRow+1
                    curItem_config=curItem_config+1

                    print "curItem_config----",curItem_config                
                    Config_check_Tree.selection_set(curItem_config)
                    time.sleep(0.5)
                    curItem_config=curItem_config+1
                    print "curItem_config----",curItem_config                
                    Config_check_Tree.selection_set(curItem_config)
    ##                time.sleep(0.5)
                    
                    
            ConfigCheck_overall_progressbar["value"]=overall_max
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(message)s')
            logging.info('##############  Config Check Completed  ##############')
            
            myAppl.Variable(Power_Supply_path).Write(129)
        except:
            logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                    level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')

            logging.exception('Test case execution stopped abrubtly')
                   
      


    #print "ITS_enabled",ITS_enabled
    if ITS_enabled == 1 :
        print "ITS_enabled",ITS_enabled
        ITS_Application_Testing()
     
    

    if Failsafe_Enabled_GUI == 1 :
     
        if Failsafe_Testing_ADAS25==1:
            Failsafe_Testing_ADAS_25()            
        else:                 
            Failsafe_Testing()
            
    if Message_Counter_enabled_GUI == 1:
        Message_Counter_Testing()

    if Active_Enabled_GUI == 1 :
        Active_Testing()        


    if Config_check_Enabled_GUI == 1 :
        Config_Check_Testing()
        
    
    if BusOff_Enabled_GUI == 1 :
        BusOff_Testing()


    if Gateway_DIAG_Enabled_GUI == 1 :
        while(Wait_Over == 0):
            time.sleep(1)
            
        Gateway_DIAG_Testing()

    if Gateway_TGW_Enabled_GUI == 1 :
        Gateway_TGW_Testing()


    if ICC_CANCEL_Enabled_GUI == 1 :
        ICC_CANCEL_Application_Testing()
        
    Summary_Sheet_Function()
    date_now =datetime.datetime.today().strftime("%Y%m%d")
    time_now=datetime.datetime.today().strftime("%H%M")
    os.rename(str(Main_Vehicle_Folder),str(Main_Vehicle_Folder+"_"+str(date_now)+"_"+(time_now)))
    Exit_Button_Function()          


    if  (Message_Counter_enabled_GUI == 0 and Gateway_DIAG_Enabled_GUI == 0 and Gateway_TGW_Enabled_GUI == 0 and Config_check_Enabled_GUI == 0 and ITS_enabled == 0):       
        print "START BUTTON pressed without selecting any button"
        #tkMessageBox.showwarning('Error:', 'No option selected')


#*********Open Log button Functionality*********#                    
def Log():
        
    os.startfile('HILS_Testing_Log.txt')

################ JSON encoder functionality for decoding dictionary ######################

def JSONTree(Tree, Parent, Dictionary):
    global uid 
    
    for key in Dictionary :
        uid = uid + 1
        if isinstance(Dictionary[key], dict):
      
            Tree.insert(Parent, 'end', uid, text=key)
            JSONTree(Tree, uid, Dictionary[key])
            Tree.item(Parent, open = True )
        elif isinstance(Dictionary[key], list):
            Tree.insert(Parent, 'end', uid, text=key + '[]', open=True)
            JSONTree(Tree , uid , dict([(i, x) for i, x in enumerate(Dictionary[key])]))
            Tree.item(Parent, open=True )
           
        else:
            value = Dictionary[key]
            if isinstance(value, str) or isinstance(value, unicode):
                value = value.replace(' ', '_')
            Tree.insert(Parent, 'end', uid, text=key, value=value)
            Tree.item(Parent, open=True)
 
##########################################################################################

############################  Construct JSON TREE ###############################################

def construct_JSON_tree(dict_Data,Frame_Number):    

    tree_frame_3 = Frame(Frame_Number, relief = RAISED,borderwidth=2, bg = frame1_color)                                                    # Creating tree frame
    tree_frame_3.place(x = 15, y = 5)
    tree_frame_3.place_configure(width = 450, height = 490)
    scrollbar = Scrollbar(tree_frame_3, bd = 3)                                                               # Creating scroll bar
    scrollbar.pack(side = RIGHT, fill = Y)
    Tree_Name = ttk.Treeview(tree_frame_3, height=25, columns= ("Testcase_name","result_col"))                                          # Creating frame for teee view
    Tree_Name.column("#0",minwidth=0,width=200, stretch=NO)
    Tree_Name.column("result_col",minwidth=0,width=50, stretch=NO)
    Tree_Name.column("Testcase_name",minwidth=0,width=150, stretch=NO)    
    Tree_Name.heading("result_col",text = "Results")
    Tree_Name.heading("Testcase_name",text = "Testcase_name")
    style1 = ttk.Style()
    style1.configure(".", font=('calibri', 11, 'bold'))
    style1.configure("Treeview", foreground='black')
    JSONTree(Tree_Name, '', dict_Data)
    scrollbar.config(command=Tree_Name.yview)
    Tree_Name.config(yscrollcommand=scrollbar.set)
    Tree_Name.pack(padx = 3,pady = 3)
    return Tree_Name

#########################################################################################
    
##################################Thread current status###################################
            
def _async_raise(tid, excobj):
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(excobj))
    if res == 0:
        raise ValueError("nonexistent thread id")
    elif res > 1:
        # """if it returns a number greater than one, you're in trouble, 
        # and you should call it again with exc=NULL to revert the effect"""
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, 0)
        raise SystemError("PyThreadState_SetAsyncExc failed")

##########################################################################################


    


##################################Thread Class ########################################### 
 
class Thread(threading.Thread):
        
#************ Thread exception *****************#
    
    def raise_exc(self, excobj):                                                                                        # Function to raise runtime exceptime in case attempt is made to stop a thread that is not started
       
        assert self.isAlive(), "thread must be started"
        for tid, tobj in threading._active.items():
            if tobj is self:
                _async_raise(tid, excobj)
                return

#***********************************************#   
        
#************ Thread termination ***************#
            
    def terminate(self):
        self.raise_exc(SystemExit)                                                                                      # "SystemExit" function used to suspend the Threads                                                                

#***********************************************#
        
##########################################################################################

##################################GUI Class############################################### 

class MyGUI(Frame):


    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI()

    def initializeUI(self):
        
        global tree_frame, tree, Script_Path, Org_Path
        global ALL_check_button,ITS_check_button,ACTIVE_check_button,MSG_COUNTER_check_button,BUSOFF_check_button,FAILSAFE_check_button,GATEWAY_check_button,CONFIG_CHECK_check_button,OK_button,ICC_Cancel_Testing_button
        global Plantmodel_button,dispatch_button,start_button,stop_button,reset_button
        global heading_color,Browse_button_color
        
        
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()

        def StartThread():
            
            global start_thread
            start_button["state"] = DISABLED
            start_thread = Thread(target = Start)
            start_thread.start()
            stop_button["state"] = NORMAL
            reset_button["state"] = DISABLED         

        

    ####################################*** Open Experiment ***######################################

        #**Open Experiment browse button Functionality**#
        def OpenExperiment():
            print "hii"
            global FilePath2, Experiment, myAppl, pfm,  VehicleName, LayoutConfig ,LayoutDiag ,LayoutMeter ,LayoutSide, LayoutEap, LayoutMrr, LayoutFrC, LayoutSow, PlatformName,platformmanager,SystemType, ApplFile, \
                   Message_Counter_Variant_Value
            myAppl=None            
            Model_Path = Org_Path + '01_Plant_Models'
            print Model_Path
            Experiment=tkFileDialog.askopenfilename(initialdir = Model_Path)
            if Experiment == '':
                tkMessageBox.showwarning('Error:', 'No file Selected')
                
            else :
                FilePath2 = os.path.dirname(Experiment)
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('reached here')

                #ADAS3_Working_set_reset\ADAS3_HILS_MODEL_L12M_B12L_L12F_L42N_151126[DS4302]'
                Plantmodel_entrybox.delete(0, END)
                Plantmodel_entrybox.insert(0, FilePath2)
                print "FilePath2",FilePath2
                print "Experiment",Experiment
                
                os.startfile(Experiment)
                

                pfm = platformmanager.Application()                                                                     # Creating an instance of the platform.
                if None == pfm.ActivePlatform:                                                                          # Checking if a platform has been set
                    raise "ERROR", "No active platform has been set. "\
                    "Please define your current platform in ControlDesk using the menu "\
                    "Platform -> Set Workingboard."                                                                     # If no active platform , then raise an exception.
                PlatformName = pfm.ActivePlatform.Name                                                                  # Collecting name of the platform set in PlatformName
                PlatformType = pfm.ActivePlatform.Type                                                                  # Collecting info whether platform is dSPACE processor (or Multiprocessor) or Simulink
                SystemType   = dSPACEDemoUtilities.GetSystemType(PlatformType)                                          # Collecting info whether dS1006 or dS1005....etc
                print "PlatformName",PlatformName
                print "SystemType",SystemType
                ApplFile   = FilePath2 + "\\ADAS_HILS_MODEL"                                                            # Specifying the name of the SDF file
                print "ApplFile",ApplFile
                if cdacon.btcSimulink == PlatformType:                                                                  # If platform set is Simulink then use offline feature of Control Desk.  
                    ApplFile = ApplFile + "_offline"
                
                # descriptions:
                if SystemType in ["MP", "MP2"]:        
                    TrcPrefix = "io_and_pid/"
                else:
                    TrcPrefix = ""
                time.sleep(2)
                myAppl = rtplib.Appl(ApplFile + ".sdf", PlatformName, SystemType)                                       # Creating an instance of Application class of rtplib
                print myAppl 
                print  'The control desk instance is created '

                if cdacon.btcSimulink == PlatformType:       
                    pfm = platformmanager.Application()                                                                 # Insert blocks into the Simulink model. Thus, if  the simulation is running, it must be stopped. 
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
                   

                Instrumentation().AnimationMode = 2
                time.sleep(1)
                logging.basicConfig(filename= 'HILS_Testing_Log.txt' , level=logging.INFO,
                                    format='%(asctime)s - %(levelname)s - %(message)s')
                logging.info('Plant model selected')
                
                
                LayoutConfig = FilePath2 + "\config_" + VehicleName +".lay"
                LayoutDiag = FilePath2 + "\diag.lay"
                LayoutMeter = FilePath2 + "\meter_navi_strg_" + VehicleName + ".lay"
            
            if(Message_Counter_enabled==1):
                MSG_COUNTER_check_button["state"] = NORMAL
                
            if (Gateway_DIAG_Enabled==1):
                GATEWAY_check_button["state"] = NORMAL
            
            if (Gateway_TGW_Enabled==1):
                GATEWAY_check_button["state"] = NORMAL

            if (BusOff_Enabled==1):
                BUSOFF_check_button["state"] = NORMAL

            if (DDT_enabled==1):
                ACTIVE_check_button["state"] = NORMAL                    
   
            if (Failsafe_Enabled==1):
                FAILSAFE_check_button["state"] = NORMAL
            
            if (Config_Check_Enabled==1):
                CONFIG_CHECK_check_button["state"]=NORMAL
                
           # Plantmodel_button["state"] = DISABLED
            reset_button["state"] = NORMAL
            
            ALL_check_button["state"] = DISABLED
            ITS_check_button["state"] = NORMAL
            FAILSAFE_check_button["state"] = NORMAL



            
            
            OK_button["state"] = DISABLED 
                  
    #***********************************************#        
    ##############################  Dispatch Sheet  ############################################################
    #*Dispatch Sheet browse and parse Functionality*#
        def Check_App_Exist(Vehicle,App,Var):                                                                       # This function will check vehicle and variant present for testcase in judgementsheet
            global Missing_Input_Details
            global All_Applications
            global error_Count 
            try:
                print "All_Applications",All_Applications
                judge_sheet_path = Org_Path + '06_Master_Judgement_Sheet\\Master_Judgement_Sheet_ITS.xls'
                check = os.path.exists(judge_sheet_path)                
                if (check == True):
                    book_JS = xlrd.open_workbook(judge_sheet_path,formatting_info=True)													# Open Judgement sheet
                    sheetNames_JS = book_JS.sheet_names()
                    sheetNumber_JS = 0
                    Sheet_No = 0
                    print "sheetNames_JS",sheetNames_JS
                    for i in sheetNames_JS:																								# Loop for finding Application sheet
                        print "i",i
                        if App == i:
                            Sheet_No = 1
                            break
                        sheetNumber_JS = sheetNumber_JS + 1

                    if Sheet_No == 1:
                        
                        App_sheet = book_JS.sheet_by_index(sheetNumber_JS)																			# variable for Py-MScript sheet
                        End_col = App_sheet.ncols
                        Application_Found = 0
                        print "Var",Var
                        for k in range(0, End_col):                
                            if App_sheet.cell(1,k).value == Vehicle:
                                print "Passed Veh",App_sheet.cell(1,k).value,End_col
                                for g in range(k, End_col):
                                    if Var in App_sheet.cell(2,g).value:
                                        print "Pass Var",App_sheet.cell(1,k).value,App_sheet.cell(2,k).value
                                        Application_Found = 1
                                        break
                                    else:
                                        continue                         
                            else:
                                continue

                        if Application_Found == 0:
                            logging.info('Info :: Vehicle not found for ' + App + ' application in Judgement sheet for '+Var)
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + ". Vehicle not found for "+App + " application in Judgement sheet for Variant "+ Var+"\n"                        
                    else:
                        
                        logging.info('Info ::' + App + ' application sheet not found in Judgement sheet for '+Var)
                        if App not in All_Applications:
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + ". " + App + ' application sheet not found in Judgement sheet \n'

                      
                else:               
                    logging.info('Info ::'  + ' Master_Judgement_Sheet_ITS not present for '+Var)
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + ". " + ' Master_Judgement_Sheet_ITS not present for '+ Var +"\n"
                    

                    #tkMessageBox.Showinfo("Info","Vehicle not found in judgement sheet")
                Book_Master_TP = Org_Path + "03_Master_Test_Sheet" + "\\" + Vehicle+"\\Master_TestPattern_ITS.xls"
                book_TP = xlrd.open_workbook(Book_Master_TP,formatting_info=True)													# Open Judgement sheet
                sheetNames_TP = book_TP.sheet_names()
                if App in sheetNames_TP and App not in All_Applications:
                    All_Applications.append(App)
                    print "Application sheet exist"
                else:
                    if App in All_Applications:
                        pass
                    else:
                        All_Applications.append(App)
                        logging.info('Info :: Application sheet not found for ' + App + ' application in Test pattern sheet')
                        error_Count = error_Count + 1
                        Missing_Input_Details += str(error_Count) + '. Application sheet not found for ' + App + ' application in Test pattern sheet\n'
            except:
                pass
                
                

        def Check_Folder_Exist(Vehicle,Failsafe_Enabled):                                                    #This function will check for vehicle folder for master test sheet and checks important keywords.
            global Missing_Input_Details,Org_Path,sig_data_sheet_str
            global Power_Supply_path,CAR_SLCT_NO_path,DIAG_CMD_NO_path,DIAG_CMD_NO_path_3,DTC_string_temp,DTC_string_1_temp,DTC_string_path_temp,Read_vehicle_speed_path
            global error_Count
            global Book_Master_TP,Test_Sheet_Path
            try:
                Folders = os.listdir(Test_Sheet_Path)
                print Folders
                #if Failsafe_Enabled == 1 :
               #     Master_TestPattern_name = 'Master_TestPattern_FLS.xls'
                
                Master_TestPattern_name = 'Master_TestPattern_ITS.xls'
                print "Master_TestPattern_name",Master_TestPattern_name
                if Vehicle in Folders:
                    logging.info('Info :: Vehicle folder found ')
                    Folders = os.listdir(Test_Sheet_Path+'\\'+Vehicle)
                    print Folders
                    if Master_TestPattern_name in Folders:
                        logging.info('Info :: Test pattern sheet found ')
                        Book_Master_TP = Org_Path + "03_Master_Test_Sheet" + "\\" + Vehicle+"\\" + Master_TestPattern_name
                        book_TP = xlrd.open_workbook(Book_Master_TP,formatting_info=True)													# Open Judgement sheet
                        sheetNames_TP = book_TP.sheet_names()
                        if sig_data_sheet_str in sheetNames_TP:
                            print "Signal Data sheet exist"
                            Signal_Data_sheet = book_TP.sheet_by_name(sig_data_sheet_str)
                            End_row = Signal_Data_sheet.nrows
                            print "End_row",End_row
                            Power_Supply_path_Found = 0
                            CAR_SLCT_NO_path_Found = 0
                            DIAG_CMD_NO_path_Found = 0
                            DIAG_CMD_NO_3_path_Found = 0
                            DTC_string_path_temp_Found = 0
                            DTC_String_path_Found = 0
                            DTC_String_1_path_Found = 0
                            Read_vehicle_speed_path_Found = 0
                            for k in range(0, End_row):
                                print Signal_Data_sheet.cell(k,0).value
                                if Signal_Data_sheet.cell(k,0).value == "Power_Supply":
                                    Power_Supply_path_Found = 1
                                    Power_Supply_path = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "CAR_SLCT_NO":
                                    CAR_SLCT_NO_path_Found = 1
                                    CAR_SLCT_NO_path = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DIAG_CMD_NO_2":
                                    DIAG_CMD_NO_path_Found = 1
                                    DIAG_CMD_NO_path = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DIAG_CMD_NO_3":
                                    DIAG_CMD_NO_3_path_Found = 1
                                    DIAG_CMD_NO_path_3 = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DTC_String":
                                    DTC_String_path_Found = 1
                                    DTC_string_temp = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DTC_String_1":
                                    DTC_String_1_path_Found = 1
                                    DTC_string_1_temp = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "DTC_string_ADAS2":
                                    DTC_string_path_temp_Found = 1
                                    DTC_string_path_temp = Signal_Data_sheet.cell(k, 1).value
                                elif Signal_Data_sheet.cell(k,0).value == "Read_vehicle_speed":
                                    Read_vehicle_speed_path_Found = 1
                                    Read_vehicle_speed_path = Signal_Data_sheet.cell(k, 1).value
                                else:
                                    continue
                                
                            if Power_Supply_path_Found == 1:
                                logging.info('Info :: Power_Supply_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: Power_Supply_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. Power_Supply_path not found in Test pattern sheet\n'
                            if CAR_SLCT_NO_path_Found == 1:
                                logging.info('Info :: CAR_SLCT_NO_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: CAR_SLCT_NO_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. CAR_SLCT_NO_path not found in Test pattern sheet\n'
                                
                            if DIAG_CMD_NO_path_Found == 1:
                                logging.info('Info :: DIAG_CMD_NO_path_2 found in Test pattern sheet')
                            else:
                                logging.info('Info :: DIAG_CMD_NO_path_2 not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. DIAG_CMD_NO_path not found in Test pattern sheet\n'
                            if DIAG_CMD_NO_3_path_Found == 1:
                                logging.info('Info :: DIAG_CMD_NO_3_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: DIAG_CMD_NO_3_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details +=str(error_Count) + '. DIAG_CMD_NO_3_path not found in Test pattern sheet\n'

                                
##                            if DTC_String_path_Found == 1:
##                                logging.info('Info :: DTC_String_path found in Test pattern sheet')
##                            else:
##                                logging.info('Info :: DTC_String_path not found in Test pattern sheet')
##                                error_Count = error_Count + 1
##                                Missing_Input_Details += str(error_Count) + '. DTC_String_path not found in Test pattern sheet\n'
##                                
##                                
##                            if DTC_String_1_path_Found == 1:
##                                logging.info('Info :: DTC_String_1_path found in Test pattern sheet')
##                            else:
##                                logging.info('Info :: DTC_String_1_path not found in Test pattern sheet')
##                                error_Count = error_Count + 1
##                                Missing_Input_Details += str(error_Count) + '. DTC_String_1_path not found in Test pattern sheet\n'
##                                
                            if DTC_string_path_temp_Found == 1:
                                logging.info('Info :: DTC_string_path_temp found in Test pattern sheet')
                            else:
                                logging.info('Info :: DTC_string_path_temp not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. DTC_string_path_temp not found in Test pattern sheet\n'
                            if Read_vehicle_speed_path_Found == 1:
                                logging.info('Info :: Read_vehicle_speed_path found in Test pattern sheet')
                            else:
                                logging.info('Info :: Read_vehicle_speed_path not found in Test pattern sheet')
                                error_Count = error_Count + 1
                                Missing_Input_Details += str(error_Count) + '. Read_vehicle_speed_path not found in Test pattern sheet\n'
                        else:
                            logging.info('Info :: Signal Data sheet not found in Test pattern sheet')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. Signal Data sheet not found in Test pattern sheet\n'
                    else:
                        print "Test pattern sheet not found"
                        logging.info('Info :: Test pattern sheet not found ')
                        error_Count = error_Count + 1
                        Missing_Input_Details += str(error_Count) + '. Test pattern sheet not found\n'
                else:                                                                                                               # if vehicle or keywords not found it will display message
                    print "Vehicle not found master test sheet folder"
                    logging.info('Info :: Vehicle not found master test sheet folder ')
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + '. Vehicle folder not found in master test sheet folder\n'
                    #tkMessageBox.Showinfo("Info","Vehicle not found in Master Test Sheet")
            except:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='folder not present')                        
                pass

        def Check_Result_Folder_Exist(App):                                             # This function will check whethere enabled applications result report folder, master result sheet and master result workbook 
            global Missing_Input_Details,All_Applications_Result,error_Count
            print "All_Applications_Result",All_Applications_Result
            try:
                Master_Result_Folder = Org_Path + "05_Master_Result_Reports"
                Folders = os.listdir(Master_Result_Folder)
                print Folders
                if App in Folders and App not in All_Applications_Result:
                    
                    logging.info('Info :: Application folder found ')
                    Folders = os.listdir(Master_Result_Folder+'\\'+App)
                    print "New",Folders
                    if App +'.xls' in Folders:                                                  #checks for master result workbook of application
                        logging.info('Info :: Result sheet found ')
                    else:
                        if App in All_Applications_Result:                                      #checks for master result sheet of application
                            print "pass"
                            pass
                        else:
                            All_Applications_Result.append(App)
                            print "Result sheet not found"
                            logging.info('Info :: Result Report not found ')
                            error_Count = error_Count + 1
                            Missing_Input_Details += str(error_Count) + '. '+App+" Master Result Report not found\n"
                else:
                    if App in All_Applications_Result:
                        pass
                    else:
                        All_Applications_Result.append(App)
                        print " Application not found master test sheet folder"
                        logging.info('Info ::  Application folder found master result report folder ')
                        error_Count = error_Count + 1
                        Missing_Input_Details += str(error_Count) + '. '+App +" folder not found in master result report folder\n"
                    #tkMessageBox.Showinfo("Info","Vehicle not found in Master Test Sheet")
            except:
                logging.basicConfig(filename= 'HILS_Testing_Log.txt',
                level=logging.INFO,format='folder not present')                        
                pass

        def DispatchSheet():
            global FilePath1, tree, tree_frame, scrollbar, Dest_Folder_Path_Vehicle,DispatchSheet,Failsafe_sheet,Failsafe_Testing_ADAS25
            global VehicleName, RegionName, Variant,  App_Arry_Final,\
                   Variant_Test_Enabled, PartNo,Variant_Value,Destination_Folder_Path,\
                   Vehicle_Id,folder_path,folder_name_app, folder_name_dispatch,AdasECU
            global Message_Counter_enabled,DDT_enabled,Failsafe_Enabled,BusOff_Enabled,Gateway_DIAG_Enabled,Gateway_TGW_Enabled,Config_Check_Enabled,DispatchSheet,ICC_Cancel_check_Enabled
            global Mot_file_name, Hils_model_name,Message_Counter_Variant_Value,GateWay_Diag_Variant_Value,GateWay_TGW_Variant_Value
            global CAN1_Busoff_Enabled,CAN2_Busoff_Enabled,CAN3_Busoff_Enabled,CAN4_Busoff_Enabled,DispatchSheet_Path_Original
            global error_Count
            global Covariant,Variant
            global Missing_Input_Details                
            global  WorkBook
            global ActiveTestDict
            global VehicleDict
            global Result_DispatchSheet_Path    
            Message_Counter_enabled=0
            DDT_enabled=0
            Gateway_TGW_Enabled=0
            Gateway_DIAG_Enabled=0
            Config_Check_Enabled = 0
            BusOff_Enabled=0
            Gateway_Enabled=0
            DispatchSheet_name=[]

            All_Process_TM = os.popen("tasklist").read()
            while "EXCEL.EXE" in All_Process_TM:
                All_Process_TM = os.popen("tasklist").read()
                os.system("taskkill /f /im EXCEL.EXE")    
            Dispatch_Path = Org_Path + '00_Dispatch_Sheets'                                                             # Path for browsing dispacth sheet 
            DispatchSheet=tkFileDialog.askopenfilename( initialdir = Dispatch_Path )                                    # Browse Dispatch File
            BUSOFF_JUDGEMENT_SHEET=  Org_Path + "\\" + "02_Script" + "\\" + "JUDGEMENT_SCRIPTS" + "\\" +"Busoff_Judgement_Sheet.xls"

            Plantmodel_button["state"] = NORMAL
            dispatch_button["state"] = DISABLED
            start_button["state"] = DISABLED
            MSG_COUNTER_check_button["state"] = DISABLED
            GATEWAY_check_button["state"] = DISABLED
            FAILSAFE_check_button["state"] = DISABLED
            OK_button["state"] = DISABLED
            
            if DispatchSheet == '':
                tkMessageBox.showwarning('Error:', 'No file Selected')
            else :
                try:
                    print 'Dispatch Sheet Selected is :'
                    print DispatchSheet
                    FilePath1 = os.path.dirname(DispatchSheet)	    														# Stores the dispacth sheet path
                    DispatchSheet_name=DispatchSheet.split('/')
                    DispatchSheet_Name=DispatchSheet_name[len(DispatchSheet_name)-1]  										# Stores dispacth sheet name
                    print "DispatchSheet_Name",str(DispatchSheet_Name)
                    print "FilePath1 is "+ str(FilePath1)
                    dispatch_entrybox.delete(0, END)																		# Clear dispatch sheet entry box contents
                    dispatch_entrybox.insert(0, DispatchSheet)																# display dispatch sheet name in entry box
                    xlapp = win32com.client.Dispatch("Excel.Application")   												# To open Excel for Message Counter Judgement Sheet
                    if os.path.exists(str(BUSOFF_JUDGEMENT_SHEET)):
                        xlapp.Workbooks.Open(Filename=str(BUSOFF_JUDGEMENT_SHEET), ReadOnly=1)
                        xlapp.Application.Run("Busoff_Judgement_Sheet.xls!module5.can_chl",str(DispatchSheet), 'Busoff_Check')
                        xlapp.Application.Quit() # Comment this out if your excel script closes
                    WorkBook =xlrd.open_workbook(DispatchSheet,
                                                     formatting_info=True)                                                      # Formatting_info=True doesnt work for xlsx
                    logging.basicConfig(filename= 'HILS_Testing_Log.txt' ,
                                            level=logging.INFO,
                                            format='%(asctime)s - %(levelname)s - %(message)s')
                    logging.info('Dispatch Sheet selected')
                    fp1 = open("HILS_Testing_Log.txt", "a")            
                    fp1.write(DispatchSheet)
                    fp1.close()
                    DSheets=WorkBook.sheet_names()
                    DIMPSheet=WorkBook.sheet_by_index(0)                                                                    # To store first sheet if dispatch sheet in var
                    DIMPSheetName=DSheets[0]                                                                                # Change index as required to get sheet name
                    DIMPSheetCol=DIMPSheet.ncols                                                                            # To find the total number of used Columns
                    DIMPSheetRow=DIMPSheet.nrows                                                                            # To find the total number of used rows
                    Failsafe_sheet=WorkBook.sheet_by_index(4)                                                               # stores fourth sheet of dispatch sheet
                except:
                    logging.info('Error ::  While Opening Dispatch sheet')
                    #*********************Gan made changes to find motfilename and Hils model starts*******************************#

                found_Hils = 0
                found_Mot = 0
##                try:
                print "try"
                for i in range(0,DIMPSheetRow):
                    for j in range(0,DIMPSheetCol):
                        try:                            
                            if str(DIMPSheet.cell(i,j).value).lower() == "mot file name":                       # Checking for the Mot File Name Keyword to get Mot File name
                                Mot_file_name = DIMPSheet.cell(i,(j + 2)).value									# Storing Mot File name to variable
                                print " Mot_file_name", Mot_file_name
                                found_Mot = 1
                                break
                        except:
                            continue
                    if(found_Mot == 1):																		# Break the loop once Mot file name found
                        break
                    
                for i in range(0,DIMPSheetRow):
                    for j in range(0,DIMPSheetCol):
                        try:
                            if("version name" in str(DIMPSheet.cell(i,j).value).lower()):									    # Checking for the Mot File Name Keyword to get version name
                                Hils_model_name = DIMPSheet.cell(i,(j + 2)).value								# Storing Mot File name to variable
                                print "Hils_model_name",Hils_model_name
                                found_Hils = 1
                                break
                        except:
                            continue
                    if(found_Hils == 1):																	# Break the loop once Version name found
                        break
                if found_Mot != 1 and found_Hils != 1:   													# If Both keywords Not Found show pop up to User
                    logging.info('Info :: The Mot file and version names not present')
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + ". Mot file and version names are not found in Dispatch sheet\n"
                    
                    #tkMessageBox.Showinfo("Info","Mot file and version names are not found in Dispatch sheet");
                elif found_Mot != 1:    												 					# If keyword Not Found show pop up to User
                    logging.info('Info ::  Mot file name Not Present')
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + ". Mot file name not found in Dispatch sheet\n"
                    #tkMessageBox.Showinfo("Info","Mot file name not found in Dispatch sheet")
                elif found_Hils != 1:																		# If keyword Not Found show pop up to User
                    logging.info('Info ::  Version name Not Present')
                    error_Count = error_Count + 1
                    Missing_Input_Details += str(error_Count) + ". version name not found in Dispatch sheet\n"

##                except:
##                    print "Error ::  While Finding the Mot file and version names"
##                    logging.info('Error ::  While Finding the Mot file and version names')
                    #*********************Gan made changes to find motfilename and Hils model ends*******************************#

                VehicleDict=OrderedDict()
                VehicleInfo = 'a'

##                xlapp = win32com.client.dynamic.Dispatch("Excel.Application")
##                xlapp.visible = False
##                own_path = os.getcwd()
##                wb = xlapp.workbooks.Open(own_path + "\\Update_Test_Folder_Structure_INI.xls")
##                xlapp.visible = False
##                time.sleep(4)
##                xlapp.Quit()
                try:
                    Destination_Folder_Path= Org_Path + '07_Result'                                                              # Path for storing the folder strucutre

                    if DIMPSheet.cell(2,2).value=='FAILSAFE_2.5':
                        Failsafe_Testing_ADAS25 =1
                    else :
                        Failsafe_Testing_ADAS25 =0
                        
                    for row in range(DIMPSheet.nrows):
                        if DIMPSheet.cell(row,2).value == 'CAN_TEST_01':
                            break
                    ItsEndRow = row + 1
                    print 'ItsEndRow', ItsEndRow
                    for i in range (0, ItsEndRow ):                                                                               # Loop for extracting the Vehicle Info
                        for j in range(0,DIMPSheetCol):
                            
                            if VehicleInfo == 'a':
                                if DIMPSheet.cell(i,j).value!='':
                                    xfx=DIMPSheet.cell_xf_index(i,j)
                                    xf=WorkBook.xf_list[xfx]
                                    pattern=xf.background.pattern_colour_index
                                    background=xf.background.background_colour_index
                                    if pattern==13 and background==64:															  # checking for the cell with specific format
                                        #print i , j 
                                        VehicleInfo=DIMPSheet.cell(i,j).value													  # Stores Vehicle information to VehicleInfo
                                        print "VehicleInfo",VehicleInfo
                                        
                            
                    PlantModel=VehicleInfo.split()																				  # Split the VehicleInfo
                    VehicleName= PlantModel[0]																					  
                    RegionName=PlantModel[1]
                    print "RegionName",RegionName
                    Vehicle_Id= VehicleName + '_' +RegionName
                    print "Vehicle_Id",Vehicle_Id
                    VehicleDict[VehicleName]= OrderedDict()				   
                    VehicleDict[VehicleName][RegionName]= OrderedDict()
                except:
                    logging.info('Error ::  While Extracing Vehicle Information')
                    
                    
                try:	
                    Dest_Folder_Path_Vehicle = Destination_Folder_Path + '\\' + Vehicle_Id                                  	  # Folder path for Vehicle Name in the folder strucutre
                    Result_DispatchSheet_Path=Dest_Folder_Path_Vehicle+"\\00_DispatchSheet_ConnectionCheck\\"+DispatchSheet_Name  # Result Folder Path for dispatch sheet
                    print "result dispatch sheet is "+str(Result_DispatchSheet_Path)
                    print " "   + str(Dest_Folder_Path_Vehicle)
                    check = os.path.isdir(Dest_Folder_Path_Vehicle)																  # checking whether vehicle directory exist or not
                    
                    if (True == check):															
                        shutil.rmtree(Dest_Folder_Path_Vehicle)																	  # If Directory exist delete it
                        time.sleep(3)
                    os.makedirs(Dest_Folder_Path_Vehicle)                                                                   	  # Creating the directory for Vehicle Name
                    folder_path = ''
                    folder_name_app = ''
                except:
                    logging.info('Error ::  While Creating Vehicle Folder')		
                try:
                    for i in range (0, ItsEndRow ):                                                                               # Loop for extracting the part number of Vehicle
                        for j in range(0,DIMPSheetCol):
                            if DIMPSheet.cell(i,j).value!='':
                                xfx=DIMPSheet.cell_xf_index(i,j)
                                xf=WorkBook.xf_list[xfx]
                                pattern=xf.background.pattern_colour_index
                                background=xf.background.background_colour_index
                                if pattern==13 and background==64:
                                    if DIMPSheet.cell(i,j).value=='Part Number':
                                        PartNo = DIMPSheet.cell(i,j+1).value                                     		          # This variant stores the Covariant's for Vehicle                                             
                                        AdasECU = DIMPSheet.cell(i,j-1).value													  # stores the Type of Ecu 
                                        if PartNo!= '':										  
                                            break
                                        else:										  
                                            continue

                    print 'ADAS'  ,  AdasECU
                except:
                    logging.info('Error ::  While Extracting Part Number')
                    
                VehicleDict[VehicleName][RegionName][PartNo]= OrderedDict()
                Covariant=[]
                Variant=[]
                ApplicationEnabled=[]
                TestEnabled=[]
                CoArray=[]
                TestCaseEnabled=[]
                var=[]
                CoRow = 0
                CoCol= 0
                try:
                    for i in range (0, ItsEndRow ):                                                                               # Loop for extracting the Variant code of Vehicle
                        for j in range(0,DIMPSheetCol):
                            if DIMPSheet.cell(i,j).value!='':
                                xfx=DIMPSheet.cell_xf_index(i,j)								
                                xf=WorkBook.xf_list[xfx]
                                pattern=xf.background.pattern_colour_index
                                background=xf.background.background_colour_index
                                if pattern==13 and background==64:
                                    if DIMPSheet.cell(i,j).value=='Variant Code':										
                                        CoRow =  i
                                        CoCol =  j
                                        
                    
                    for j in range(CoCol + 1,DIMPSheetCol):                                                                 # Loop for extracting the Covariant of Vehicle
                            if DIMPSheet.cell(CoRow,j).value!='':
                                xfx=DIMPSheet.cell_xf_index(CoRow,j)
                                xf=WorkBook.xf_list[xfx]
                                pattern=xf.background.pattern_colour_index
                                background=xf.background.background_colour_index
                                if pattern==13 and background==64:
                                    Covariant.append(DIMPSheet.cell(CoRow,j).value)                                         # This array stores the Covariant's for Vehicle 
                                    Variant_Value.append(DIMPSheet.cell(CoRow+1,j).value)
                                    CoArray.append(j)
                    for i in range(0,len(Covariant)):                                                                       
                        Variant.append(PartNo+ '_' + str(int(Variant_Value[i])))                                            # Appending the Variant_Value to part number                                                                                                                                      # Variant array contains the variant name (Variant_Value and part number)
                    for variants in Variant:
                        VehicleDict[VehicleName][RegionName][PartNo][variants]=OrderedDict()
                    print "Covariant",Covariant
                    if len(Covariant) == 0:																					# If None Of the variants enabled show message to user
                        logging.info('Info :: None of the variants enabled')
                        error_Count = error_Count + 1
                        Missing_Input_Details += str(error_Count) + ". None of the variants enabled in Dispatch sheet\n"
                        #tkMessageBox.Showinfo("Info","None of the variants enabled in Dispatch sheet")
                except:
                    logging.info('Error ::  While Extracting Covariants')
                App_Arry = []
                App_Arry_Final = []
                testNo_cnt_Final=[]
                print 'Variant_Value', Variant_Value
                try:
                    for i in range (0 ,ItsEndRow):
                        for j in range(0,DIMPSheetCol):
                            if DIMPSheet.cell(i,j).value == 'Application':                                                  # Searching "Application" keyword in Dispatch Sheet"                                              
                                    ApplicationCol = j
                                    ApplicationRow = i
                                    if ApplicationCol!= '':
                                        break
                                    else:
                                        continue

                    Appl_Path = []
                    Active_Test_Case_Arr=[]
                    Active_Test_Case_Arr_Final = []
                    for i in range(9,ItsEndRow):
                        
                        if DIMPSheet.cell(i, ApplicationCol).value != '':
                            val = DIMPSheet.cell(i, ApplicationCol).value 
                            Appl_Path.append(val)                                                                           # Extracting all application names from dispatch sheet
                    
                    Var_Path_Test_Final_arr = []
                    App_Arry_Indi = []
                except:
                    logging.info('Error ::  While Extracting Applications from DispatchSheet')
               
                for i in range (0, ItsEndRow ):                                                                               # Searching "TestCase ID" keyword in Dispatch Sheet"                                                                                                                     
                    for j in range(0,DIMPSheetCol):
                        if DIMPSheet.cell(i,j).value!='':
                            xfx=DIMPSheet.cell_xf_index(i,j)
                            xf=WorkBook.xf_list[xfx]
                            pattern=xf.background.pattern_colour_index
                            background=xf.background.background_colour_index
                            if pattern==13 and background==64:
                                if DIMPSheet.cell(i,j).value=='TestCase ID':
                                    TestCaseIdCol = j
                                    TestCaseIdRow = i
                                    if TestCaseIdCol!= '':
                                        break
                                    else:

                                        continue

                print "TestCaseIdCol",TestCaseIdCol,TestCaseIdRow               
                for j in range(0,len(Variant)):
                    print "j",j,len(Variant)
                    app_cnt = 0
                    testNo_cnt = []                    
                    for i in range (TestCaseIdRow,ItsEndRow):
                        if DIMPSheet.cell(i,CoArray[j]).value!='':
                            xfx=DIMPSheet.cell_xf_index(i,CoArray[j])
                            xf=WorkBook.xf_list[xfx]
                            pattern=xf.background.pattern_colour_index
                            background=xf.background.background_colour_index
                            if pattern == 9 and background == 64:
                                val = DIMPSheet.cell(i, TestCaseIdCol).value
                                print "val",val,TestCaseIdCol                           
                                Active_Test_Case_Arr.append(val)                                                        # This array contains all active test cases 
                                s1 = 'TEST'
                                ind = val.index(s1)
                               
                                val1 = val[0:ind-1]
                                if ((val1) in App_Arry):
                                    app_cnt= app_cnt + 1
                                    
                                else :  
                                    App_Arry.append(val1)
                                    testNo_cnt.append(app_cnt)
                                    app_cnt = 1
                            
                    Active_Test_Case_Arr_Final.append(Active_Test_Case_Arr)
                    testNo_cnt.append(app_cnt)              
                    del testNo_cnt[0]
                    testNo_cnt_Final.append(testNo_cnt)                                                                 # This array contains number of test cases in each active application for each variant      
                    App_Arry_Final.append(App_Arry)                                                                     # This array contains the active applications for each variant
                    App_Arry=[]


                
                for j in range(0,len(Variant)):
                    for i in range(0,len(App_Arry_Final[j])):
                        VehicleDict[VehicleName][RegionName][PartNo]\
                        [Variant[j]][App_Arry_Final[j][i]]=OrderedDict()

                                          

                Variant_Test_Enabled=[]

                for j in CoArray:
                    #print "TestCaseIdRow",TestCaseIdRow
                    for i in range (TestCaseIdRow,ItsEndRow):
                        if DIMPSheet.cell(i,j).value!='':
                            xfx=DIMPSheet.cell_xf_index(i,j)
                            xf=WorkBook.xf_list[xfx]
                            pattern=xf.background.pattern_colour_index
                            background=xf.background.background_colour_index
                            if pattern==9 and background==64:
                                TestCaseEnabled.append(DIMPSheet.cell(i,TestCaseIdCol).value)

                                print "TestCaseEnabled",TestCaseEnabled
                    Variant_Test_Enabled.append(TestCaseEnabled)                                                        # This array contains active test cases for each variant
                    TestCaseEnabled= []

            
                

                    
                Dest_Folder_Path_Level_array = []



#**************************  Script to check DISPATCH SHEET for enabled APPLICATIONS like Message Counter Testing, DDT Testing, Gateway Testing, Failsafe , Bus Off  ************************#
                
                global active_test_cases_column
                global active_test_cases_row
                global message_counter_column,col_arr
                global message_counter_row
                global Config_check_row,Config_check_col,Config_check_col_new
                global  Failsafe_Enabled, Message_Counter_enabled, Gateway_TGW_Enabled,Config_Check_Enabled,Number_of_BusOff_CAN_Channel,ICC_Cancel_Check_Enabled
                global Config_check_enabled_tree,Config_check_enabled_tree_col
                Config_check_enabled_tree=[]
                Config_check_enabled_tree_col=[]
                col_arr=[]
                enabled_active_list = []
                Number_of_BusOff_CAN_Channel=[]

                def cell_validation(x,y):
                
                    if DIMPSheet.cell(x,y+1).value!='':
                       
                        xfx=DIMPSheet.cell_xf_index(x,y+1)
                        xf=WorkBook.xf_list[xfx]
                        pattern=xf.background.pattern_colour_index
                        background=xf.background.background_colour_index
                        if pattern==9 and background==64:
                          return 1
                        else:
                            return 0
                        
                            
                    
                
                for i in range (0, DIMPSheetRow ):                                                                               # Searching "TestCase ID" keyword in Dispatch Sheet"                                                                                                                     
                    for j in range(0,DIMPSheetCol):
                        if DIMPSheet.cell(i,j).value=='ACTIVE TEST CASES':                                     #this finds row and column of "Message Counter" in main dispatch sheet
                            active_test_cases_row=i
                            active_test_cases_column=j

                        if DIMPSheet.cell(i,j).value=='Variant Code':
                                    
                            CoRow =  i
                            CoCol= j
                                    
                        for j in range(CoCol + 1,DIMPSheetCol):                                                                 # Loop for extracting the Covariant of Vehicle
                            if DIMPSheet.cell(CoRow,j).value!='':
                                
                                xfx=DIMPSheet.cell_xf_index(CoRow,j)
                                xf=WorkBook.xf_list[xfx]
                                pattern=xf.background.pattern_colour_index
                                background=xf.background.background_colour_index

                                if pattern==13 and background==64:
                    
                                    col_arr.append(j)
                                    
                                      


                for i in range (0, DIMPSheetRow ):                                                                               # Searching "TestCase ID" keyword in Dispatch Sheet"                                                                                                                     
                    for j in range(0,DIMPSheetCol):                                
                        if DIMPSheet.cell(i,j).value=='Config_check':                                         #This will find "Config_check" in dispatch sheet
                            print"Config_check found in dispatch "
                            Config_Check_Enabled =1
                            Config_check_row = i

                            for k in range(0,len(Variant)):

                                if DIMPSheet.cell(Config_check_row,col_arr[k]).value!= '' :
                                    xfx=DIMPSheet.cell_xf_index(Config_check_row,col_arr[k])
                            
                                    xf=WorkBook.xf_list[xfx]
                                    pattern=xf.background.pattern_colour_index
                                    background=xf.background.background_colour_index
                                    if pattern==9 and background==64:                                               #This will find which one is enabled 
                                        
                                        Config_check_enabled_tree.append('Config_check_Testing')
                                        print "Config_check_enabled_tree",Config_check_enabled_tree
                                        
                                    else :
                                        Config_check_enabled_tree.append('')
                Message_Counter_Variant_Value = []
                GateWay_Diag_Variant_Value = []
                GateWay_TGW_Variant_Value = []
                for i in range(active_test_cases_row+1,active_test_cases_row+9):
                    if DIMPSheet.cell(i,active_test_cases_column).value=='Message Counter':
                        if cell_validation(i , active_test_cases_column)==1:
                            Message_Counter_enabled=1
                            enabled_active_list.append("Message Counter")
                            for k in range(0,len(Variant_Value)):
                                if DIMPSheet.cell(i,col_arr[k]).value!= '' :
                                    xfx=DIMPSheet.cell_xf_index(i,col_arr[k])                                
                                    xf=WorkBook.xf_list[xfx]
                                    pattern=xf.background.pattern_colour_index
                                    background=xf.background.background_colour_index
                                    if pattern==9 and background==64:                                               #This will find which one is enabled                                                                                         
                                        Message_Counter_Variant_Value.append(Variant_Value[k])
                                    else :
                                        print "Message_Counter_enable Not "                                            
                        else:
                            Message_Counter_enabled=0                    
                    elif DIMPSheet.cell(i,active_test_cases_column).value=='Gateway_DIAG':
                        if cell_validation(i , active_test_cases_column)==1:
                            Gateway_DIAG_Enabled=1
                            enabled_active_list.append("Gateway_DIAG")                                
                            for k in range(0,len(Variant_Value)):
                                if DIMPSheet.cell(i,col_arr[k]).value!= '' :
                                    xfx=DIMPSheet.cell_xf_index(i,col_arr[k])                                
                                    xf=WorkBook.xf_list[xfx]
                                    pattern=xf.background.pattern_colour_index
                                    background=xf.background.background_colour_index
                                    if pattern==9 and background==64:                                               #This will find which one is enabled                                                                                         
                                        GateWay_Diag_Variant_Value.append(Variant_Value[k])
                                    else :
                                        print "GateWayDiag Not Enable"                                            
                        else:
                            Gateway_DIAG_Enabled=0
                    

                    elif DIMPSheet.cell(i,active_test_cases_column).value=='ICC_Cancel_Check':
                        if cell_validation(i , active_test_cases_column)==1:
                            ICC_Cancel_Check_Enabled=1
                            enabled_active_list.append("ICC_Cancel_Check")                                                                           
                        else:
                            ICC_Cancel_Check_Enabled=0
                    
                            
                                                        
                            
                    elif DIMPSheet.cell(i,active_test_cases_column).value=='Gateway_TGW':
                        if cell_validation(i , active_test_cases_column)==1:
                            Gateway_TGW_Enabled=1
                            enabled_active_list.append("Gateway_TGW")
                            for k in range(0,len(Variant_Value)):
                                if DIMPSheet.cell(i,col_arr[k]).value!= '' :
                                    xfx=DIMPSheet.cell_xf_index(i,col_arr[k])                                
                                    xf=WorkBook.xf_list[xfx]
                                    pattern=xf.background.pattern_colour_index
                                    background=xf.background.background_colour_index
                                    if pattern==9 and background==64:                                               #This will find which one is enabled                                                                                         
                                        GateWay_TGW_Variant_Value.append(Variant_Value[k])
                                    else :
                                        print "GateWayTGW Not Enable"                                            
                        else:
                            Gateway_TGW_Enabled=0 
                               
                    elif DIMPSheet.cell(i,active_test_cases_column).value=='Failsafe Table':
                   
                     
                        if cell_validation(i , active_test_cases_column)==1:                      
                                Failsafe_Enabled=1
                                enabled_active_list.append("Failsafe Table")
                        else:
                            Failsafe_Enabled=0
                                
                    elif DIMPSheet.cell(i,active_test_cases_column).value=='CAN1 Busoff':
                        if cell_validation(i , active_test_cases_column)==1:
                                BusOff_Enabled=1
                                CAN1_Busoff_Enabled=1
                                enabled_active_list.append("Bus off recovery")
                                Number_of_BusOff_CAN_Channel.append("CAN 1")

                    elif DIMPSheet.cell(i,active_test_cases_column).value=='CAN2 Busoff':
                        if cell_validation(i , active_test_cases_column)==1:
                                BusOff_Enabled=1
                                CAN2_Busoff_Enabled=1
                                enabled_active_list.append("Bus off recovery")
                                Number_of_BusOff_CAN_Channel.append("CAN 2")

                    elif DIMPSheet.cell(i,active_test_cases_column).value=='CAN3 Busoff':
                        if cell_validation(i , active_test_cases_column)==1:
                                BusOff_Enabled=1
                                CAN3_Busoff_Enabled=1
                                Number_of_BusOff_CAN_Channel.append("CAN 3")

                    elif DIMPSheet.cell(i,active_test_cases_column).value=='CAN4 Busoff':
                        if cell_validation(i , active_test_cases_column)==1:
                                BusOff_Enabled=1
                                CAN4_Busoff_Enabled=1
                                enabled_active_list.append("Bus off recovery")
                                Number_of_BusOff_CAN_Channel.append("CAN 4")
                       # else:
                       #      Busoff_Enabled=0
                    else:
                        print "No application enabled"

                print enabled_active_list
                for i in range (0, DIMPSheetRow ):                                                                               # Searching "TestCase ID" keyword in Dispatch Sheet"                                                                                                                     
                    for j in range(0,DIMPSheetCol):
                        if DIMPSheet.cell(i,j).value=='DDT TEST CASES':                                     #this finds row and column of "Message Counter" in main dispatch sheet
                            ddt_test_cases_row=i
                            ddt_test_cases_column=j
                            if DIMPSheet.cell(ddt_test_cases_row,ddt_test_cases_column+1).value!='':
                       
                                xfx=DIMPSheet.cell_xf_index(ddt_test_cases_row,ddt_test_cases_column+1)
                                xf=WorkBook.xf_list[xfx]
                                pattern=xf.background.pattern_colour_index
                                background=xf.background.background_colour_index
                                if pattern==13 and background==64:
                                    DDT_enabled=1
                                else :
                                    DDT_enabled=0




                k = 0
                print "App_Arry_Final",App_Arry_Final
                Check_Folder_Exist(VehicleName,Failsafe_Enabled)
                

             
                for Var_Index in range(0,len(Variant)):
                    for App_Index in range (0,len(App_Arry_Final[Var_Index])):
                        print "Variant[Var_Index]",Variant[Var_Index]
                        Check_App_Exist(VehicleName,App_Arry_Final[Var_Index][App_Index],RegionName+"_"+Variant[Var_Index].rsplit("_",1)[1])
                        Check_Result_Folder_Exist(App_Arry_Final[Var_Index][App_Index])



                if Missing_Input_Details == "":
                    pass
                else:
                    tkMessageBox.showwarning("Info",Missing_Input_Details)                                      # show all messages of user mistakes and exit code
##                    All_Process_TM = os.popen("tasklist").read()
##                    while "EXCEL.EXE" in All_Process_TM:
##                        All_Process_TM = os.popen("tasklist").read()
##                        os.system("taskkill /f /im EXCEL.EXE") 
##                    cmd = 'WMIC PROCESS get Caption,Commandline,Processid'
##                    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)    
##                    for task in proc.stdout:        
##                        if "Pythonwin.exe" in task:
##                            os.system('TASKKILL /F /IM PythonWin.exe')   
##                          
#**************************************************************************************************************************************************************************#
                                                               
#********************************************************************************************************************************************************************************************#                         
                
                
                def Safe_Make_Folder(i):                                                                                # Function definition for for folder creation
                
                   try:  
                        os.mkdir(i)
                   except:
                        pass
                if Failsafe_Enabled==1 and Failsafe_Testing_ADAS25==1:
                    Folder_Structure_Sheet_Path = Script_Path + "\\" + "Failsafe_Test_Folder_Structure_INI.xls"

                else:
                    Folder_Structure_Sheet_Path = Script_Path + "\\" + "Test_Folder_Structure_INI.xls"
                    
                    
                # Creating path for excel sheet which contains folder structure
                
                WorkBook_folder = xlrd.open_workbook(Folder_Structure_Sheet_Path                                        # Opening and Extracting information of folder strucutre excel sheet
                                                     ,formatting_info=True)
                DSheets_Fs=WorkBook_folder.sheet_names()
                FolderSheet_Fs=WorkBook_folder.sheet_by_index(0)
                DIMPSheetCol_Fs=FolderSheet_Fs.ncols                                                                    # Number of columns
                DIMPSheetRow_Fs = FolderSheet_Fs.nrows                                                                  # Number of rows 


                i_prev=0
                i_prev_array = []
               
                Dest_Folder_Path_Level_prev = ''
                
                Dest_Folder_Path_Level = ''
                Dest_Folder_Path_Level_prev_arr = []
                Dest_Folder_Path_level = Dest_Folder_Path_Vehicle     
                Dest_Folder_Path_Level_prev = Dest_Folder_Path_level
                count = 0
                temp_i_prev_array = []
                temp_Dest_Folder_Path_Level_prev_arr = []



                for i in range (i_prev ,DIMPSheetRow_Fs-1):                                                             # Loop for extracting "Applications" keyword from folder structure creation excel sheet for appending
                    for j in range (0,DIMPSheetCol_Fs-1):
                        print FolderSheet_Fs.cell(i,j).value ,"a"
                        if("Applications" in FolderSheet_Fs.cell(i,j).value):
                            folder_name_app = FolderSheet_Fs.cell(i,j).value
                            
                            break
                        
                for i in range (i_prev ,DIMPSheetRow_Fs-1):                                                             # Loop for extracting "Applications" keyword from folder structure creation excel sheet for appending
                    for j in range (0,DIMPSheetCol_Fs-1):
                        if("DispatchSheet" in FolderSheet_Fs.cell(i,j).value):
                            folder_name_dispatch = FolderSheet_Fs.cell(i,j).value
                           
                            break
                        
                
                for j in range (0,DIMPSheetCol_Fs-1):

                    for i in range (i_prev ,DIMPSheetRow_Fs-1):
                        if(FolderSheet_Fs.cell(i,j).value != ''):
                            print "if",FolderSheet_Fs.cell(i,j).value
                            for k in range (0,len(temp_i_prev_array)):
                                if i > temp_i_prev_array[k]:
                                    
                                    continue
                                
                                elif i == temp_i_prev_array[k]:
                                    count = k
                                    Dest_Folder_Path_Level_prev = temp_Dest_Folder_Path_Level_prev_arr[count]
                                    break
                                    
                                else :
                                    count = k-1
                                    break 
                            Dest_Folder_Path_Level = Dest_Folder_Path_Level_prev + "\\" + FolderSheet_Fs.cell(i,j).value
                            
                            Safe_Make_Folder(Dest_Folder_Path_Level)
                            Dest_Folder_Path_Level_array.append(Dest_Folder_Path_Level)                                 # This array contains all paths of folder structure
                            
                        elif (FolderSheet_Fs.cell(i,j).value == '' and  FolderSheet_Fs.cell(i-1,j).value != ''):
                            print "FolderSheet_Fs.cell(i,j).value",FolderSheet_Fs.cell(i-1,j).value
                            i_prev_array.append(i)
                            Dest_Folder_Path_Level_prev_arr.append(Dest_Folder_Path_Level)
                            continue
                        else :
                            continue 

                    i_prev= i_prev_array[0]              
                    Dest_Folder_Path_Level_prev = Dest_Folder_Path_Level_prev_arr[count]
                    temp_i_prev_array = i_prev_array
                    temp_Dest_Folder_Path_Level_prev_arr = Dest_Folder_Path_Level_prev_arr
                    
                    i_prev_array = []
                    Dest_Folder_Path_Level_prev_arr = []
                        
                folder_copy_path_dest = ''
                
                            
                for i in range(0,len(Dest_Folder_Path_Level_array)):                                                    # Loop for extracting the path of Vehicle Variants in Result folder"
                    if "Vehicle Variants" in Dest_Folder_Path_Level_array[i]:
                        folder_copy_path_src = Dest_Folder_Path_Level_array[i]
                        break
                s = len(folder_copy_path_src) - 19
                print "folder_copy_path_src",folder_copy_path_src
                folder_copy_src = folder_copy_path_src[:s]                                                              # Extracting the base path of Vehicle Variants in Result folder which is used for rename

                folder_copy_path_dest_arr = []
                folder_copy_path_arr = []
                
                folder_path = folder_copy_path_src[:(s-1)]
                print "folder_path",folder_path
                folder_copy_path_dest_arr.append(folder_copy_path_src)
                
                for i in range(0,(len(Variant)-1)):
                    folder_copy_path_dest =  folder_copy_path_src + str(i)
                    shutil.copytree(folder_copy_path_src, folder_copy_path_dest)
                    folder_copy_path_arr.append(folder_copy_path_dest)
                    folder_copy_path_dest_arr.append(folder_copy_path_dest)                                             # This array contains the Vehicle Variant path before renaming
                global VehicleNameFolder
                global Main_Vehicle_Folder
                VehicleNameFolder = []
                print Vehicle_Id, "Vehicle_Id"
                for i in range(0,len(Variant)):
                    Main_Vehicle_Folder= Dest_Folder_Path_Vehicle
                    print "Main_Vehicle_Folder",Main_Vehicle_Folder
                    VehicleNameFol = folder_copy_src + Vehicle_Id + "_" + str(int(Variant_Value[i]))
                    VehicleNameFolder.append(VehicleNameFol)                                                            # This array contains complete path of Vehicle variant in Result folder
                    
                for i in range(0,len(folder_copy_path_dest_arr)):                                                      # Loop for renaming Vehicle variant folders in Result folder
                    print "folder_copy_path_dest_arr",folder_copy_path_dest_arr[i]
                    print "VehicleNameFolder",VehicleNameFolder[i]
                    os.rename(folder_copy_path_dest_arr[i],VehicleNameFolder[i])

                
                for j in range(0,len(Variant)):
                    h=0
                    for i in range(0,len(App_Arry_Final[j])):
                        for k in range((testNo_cnt_Final[j][i]),0,-1):                                                  # Creating a tree structure of dictionary data on tree frame
                            VehicleDict[VehicleName][RegionName]\
                            [PartNo][Variant[j]][App_Arry_Final[j][i]]\
                            [Variant_Test_Enabled[j][h]]=OrderedDict()
                            h= h+1
                          



                
#**************************************************************       ACTIVE TEST CASE SCRIPT            ********************************************************************************************#                   
           
          
            
                ActiveTestDict = OrderedDict()
                ActiveTestDict[VehicleName]= OrderedDict()
                ActiveTestDict[VehicleName][RegionName]= OrderedDict()
                ActiveTestDict[VehicleName][RegionName][PartNo]= OrderedDict()
                Active_Test = 'ACTIVE TESTING'
                ActiveTestDict[VehicleName][RegionName][PartNo][Active_Test]=OrderedDict()
                Active_Test_Case_Types= ("DDT TESTING","MSG COUNTER","GATEWAY","BUS OFF")
                
                for k in range(0,len(Active_Test_Case_Types )):
                    ActiveTestDict[VehicleName][RegionName][PartNo][Active_Test][Active_Test_Case_Types[k]]=OrderedDict()

                #print ActiveTestDict
                     
                    

#*******************************************************************  END OF ACTIVE TEST CASE SCRIPT  ********************************************************************************#        


        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        background_frame.pack(fill=BOTH, expand=True)

        heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        heading_frame = Frame(background_frame,relief=RAISED,borderwidth=3,
                              bg = heading_color)                                                                       # Creating and Setting the design of Heading Frame
        heading_frame.place(x=15, y=15)
        heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(heading_frame, image = Nissan_image, bg = heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(heading_frame, image = TCS_image, bg = heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        
         
     
        browse_frame = Frame(background_frame, relief = RAISED,                                                                 
                             borderwidth=2, bg = frame1_color)                                                          # Creating and setting the design of dispatch and plant model frame         
        browse_frame.place(x = 15, y = 80)
        browse_frame.place_configure(width = 450, height = 180)


        
       
        dispatch_frame = Frame(browse_frame, relief = RAISED,
                               borderwidth=2, bg = browse_frame_color)
        dispatch_frame.place(x = 10, y = 3)
        dispatch_frame.place_configure(width = 425, height = 80)

        buttons_frame = Frame (background_frame, relief = RAISED,
                               borderwidth = 2)                                                      # Creating a frame for start, stop and reset button
        buttons_frame.place(x = 15, y = 265)

        buttons_frame.place_configure(width = 450, height = 60)        
        heading_label= Label(heading_frame, justify="center",
                             text = "ADAS HILS - AUTOMATION FRAMEWORK",                                                 # Creating the Heading Label
                             bg =heading_color, fg = "black", font="Times 25 bold")
        heading_label.pack()

       
        dispatch_label = Label(dispatch_frame,
                               text = "Vehicle Dispatch Sheet Selection",
                               bg = browse_frame_color , fg = "black",
                               font =("times",13,'bold'))                                                                        # Creating and setting dispatch frame
        dispatch_label.pack()

        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)
        dispatch_button = Tkinter.Button(dispatch_frame,
                                         text = "Browse",
                                         bg = Browse_button_color,
                                         activebackground = "red",
                                         height = 2, width = 8, relief = RAISED,
                                         bd = 3, cursor = "hand2", command= DispatchSheet,
                                         font = ("arial",10,"bold"))
        dispatch_button.pack(side = RIGHT, padx = 5, pady =4)

        dispatch_entrybox =Entry(dispatch_frame, width = 70, bd = 5)
        dispatch_entrybox.pack(side = LEFT, pady = 6, padx = 4)
        dispatch_entrybox.insert(0, FilePath1)

        dispatch_folder_logo = Image.open(Image_Path+"Folder.png")                                                                # Adding folder image to the dispatch sheet browsing functionality
        dispatch_folder_logo = dispatch_folder_logo.resize((30, 30), Image.ANTIALIAS)
        dispatch_folder_image = ImageTk.PhotoImage(dispatch_folder_logo)
        f1_label = Label(dispatch_frame, image = dispatch_folder_image, bg = browse_frame_color)
        f1_label.image = dispatch_folder_image
        f1_label.place(x = 0, y = 0) 

        
        Plantmodel_frame = Frame(browse_frame,
                                 relief = RAISED,
                                 borderwidth=2, bg = browse_frame_color)                                                # Creating and setting plant model frame                  
        Plantmodel_frame.place(x = 10, y = 90)
        Plantmodel_frame.place_configure(width = 425, height = 80)        

        Plantmodel_label = Label(Plantmodel_frame,
                                 text = "Vehicle Plant Model Selection",
                                 bg = browse_frame_color , fg = "black",
                                 font =("times",13,'bold'))                                                                    
        Plantmodel_label.pack()

        Plantmodel_button = Tkinter.Button(Plantmodel_frame,
                                           text = "Browse", bg = Browse_button_color,
                                           activebackground = "red", height = 2,
                                           width = 8,
                                           relief = RAISED, bd = 3, cursor = "hand2",
                                           command=OpenExperiment,
                                           font = ("arial",10,"bold"))
        Plantmodel_button.pack(side = RIGHT, padx = 5, pady =4)

        Plantmodel_entrybox = Entry(Plantmodel_frame, width = 70, bd = 5)
        Plantmodel_entrybox.pack(side = LEFT, pady = 6, padx = 4)
        Plantmodel_entrybox.insert(0, FilePath2)

        Plantmodel_folder_logo = Image.open(Image_Path+"Folder.png")                                                                # Adding folder image to the Plant model browsing functionality
        Plantmodel_folder_logo = Plantmodel_folder_logo.resize((30, 30), Image.ANTIALIAS)
        Plantmodel_folder_image = ImageTk.PhotoImage(Plantmodel_folder_logo)
        f1_label = Label(Plantmodel_frame, image = Plantmodel_folder_image, bg = browse_frame_color)
        f1_label.image = Plantmodel_folder_image
        f1_label.place(x = 0, y = 0)            

      
        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 80)

        start_button_color = '#%02x%02x%02x' % (146, 208, 80)
        start_button = Tkinter.Button(buttons_frame, text = "START",
                                      bg = start_button_color, width = 10,
                                      height = 3, relief = GROOVE,
                                      activebackground = "red", cursor = "hand2",
                                      font = "times 10 bold", command= StartThread)
        start_button.pack(side = LEFT, padx = 10, pady = 2)
        stop_button_color = '#%02x%02x%02x' % (197, 90, 17)
        stop_button = Tkinter.Button(buttons_frame, text = "STOP",
                                     bg = stop_button_color, width = 10, height = 3,
                                     relief = GROOVE, activebackground = "red",
                                     cursor = "hand2",font = "times 10 bold",
                                     command=Stop)
        stop_button.pack(side = LEFT, padx = 80, pady = 2)
         #Reset button    
        reset_button_color = '#%02x%02x%02x' % (113, 166, 219)
        reset_button = Tkinter.Button(buttons_frame, text = "CLEAR",
                                      bg = reset_button_color, width = 10,
                                      height = 3, relief = GROOVE,
                                      activebackground = "red", cursor = "hand2",
                                      font = "times 10 bold", command = Reset)
        reset_button.pack(side = RIGHT, padx = 20, pady = 2)

        
        
      

        frame3 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame3.pack(fill = BOTH, expand = TRUE)
        logging_frame = LabelFrame(frame3, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =145, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame

        s = ttk.Style()
        s.theme_use('default')
        s.configure("green.Horizontal.TProgressbar", troughcolor  ="Dark Slate Gray")


 
  
        global v1,v2,v3,v4,v5,v6,v7,v8,v9,v10

        v1=IntVar()     # ALL button
        v2=IntVar()     #ITS button
        v3=IntVar()     #DDT button
        v4=IntVar()     #MSG_COUNTER button
        v5=IntVar()     #FAILSAFE button
        v6=IntVar()     #GATEWAY button
        v7=IntVar()     #BUSOFF button
        v8=IntVar()     #CONFIG CHECK button
        v9=IntVar()     #Maneuver Testing Button
        v10=IntVar()    # Active_testing



#####################################Frames for Sequence nume#############################################################################################    

        ITS_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        ITS_frame_seq.place(x=20 ,y=20)
        ITS_frame_seq.place_configure(width = 60, height = 70)
        ITS_frame_seq_label= Label(ITS_frame_seq, justify="center",
                             text = "1.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        ITS_frame_seq_label.place(x=12,y=4)
        
        FAILSAFE_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        FAILSAFE_frame_seq.place(x=20 ,y=100)
        FAILSAFE_frame_seq.place_configure(width = 60, height = 70)
        FAILSAFE_frame_seq_label= Label(FAILSAFE_frame_seq, justify="center",
                             text = "2.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")

        FAILSAFE_frame_seq_label.place(x=12,y=4)

        MSG_CNTR_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        MSG_CNTR_frame_seq.place(x=20 ,y=180)
        MSG_CNTR_frame_seq.place_configure(width = 60, height = 70)
        MSG_CNTR_frame_seq_label= Label(MSG_CNTR_frame_seq, justify="center",
                             text = "3.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        MSG_CNTR_frame_seq_label.place(x=12,y=4)

        CONFIG_CHECK_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        CONFIG_CHECK_frame_seq.place(x=20 ,y=260)
        CONFIG_CHECK_frame_seq.place_configure(width = 60, height = 70)
        CONFIG_CHECK_frame_seq_label= Label(CONFIG_CHECK_frame_seq, justify="center",
                             text = "4.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        CONFIG_CHECK_frame_seq_label.place(x=12,y=5)


        GATEWAY_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        GATEWAY_frame_seq.place(x=550 ,y=180)
        GATEWAY_frame_seq.place_configure(width = 60, height = 70)
        GATEWAY_frame_seq_label= Label(GATEWAY_frame_seq, justify="center",
                             text = "7.",                                                 # Creating the Heading Label
                             fg = "black", font="Times 25 bold")
        GATEWAY_frame_seq_label.place(x=12,y=4)

        BUSOFF_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        BUSOFF_frame_seq.place(x=550 ,y=20)
        BUSOFF_frame_seq.place_configure(width = 60, height = 70)
        BUSOFF_frame_seq_label= Label(BUSOFF_frame_seq, justify="center",
                             text = "5.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        BUSOFF_frame_seq_label.place(x=12,y=4)
        
        ACTIVE_frame_seq= Frame(frame3,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
        ACTIVE_frame_seq.place(x=550 ,y=260)
        ACTIVE_frame_seq.place_configure(width = 60, height = 70)
        ACTIVE_frame_seq_label= Label(ACTIVE_frame_seq, justify="center",
                             text = "8.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        ACTIVE_frame_seq_label.place(x=12,y=4)

##        Maneuver_Testing_frame_seq = Frame(frame3,self,relief = RAISED, borderwidth = 4,bg =frame1_color, height = 20)                                                              # Create a new frame for ITS button 
##        Maneuver_Testing_frame_seq.place(x=550 ,y=100)
##        Maneuver_Testing_frame_seq.place_configure(width = 60, height = 70)
##        Maneuver_Testing_frame_seq_label= Label(Maneuver_Testing_frame_seq, justify="center",
##                             text = "6.",                                                 # Creating the Heading Label
##                              fg = "black", font="Times 25 bold")
##        Maneuver_Testing_frame_seq_label.place(x=12,y=4)         

        ICC_cancel_Testing_frame_seq = Frame(frame3,self,relief = RAISED, borderwidth = 4,bg =frame1_color, height = 10)                                                              # Create a new frame for ITS button 
        ICC_cancel_Testing_frame_seq.place(x=550 ,y=100)
        ICC_cancel_Testing_frame_seq.place_configure(width = 60, height = 70)
        ICC_cancel_Testing_frame_seq_label= Label(ICC_cancel_Testing_frame_seq, justify="center",
                             text = "6.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        ICC_cancel_Testing_frame_seq_label.place(x=12,y=4) 
        
#####################################Frames for CheckButtons#############################################################################################            
        ALL_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for all button 
        ALL_frame.place(x=380 ,y=340)
        ALL_frame.place_configure(width = 390, height = 70)
        
        ITS_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for ITS button 
        ITS_frame.place(x=80 ,y=20)
        ITS_frame.place_configure(width = 390, height = 70)

        ACTIVE_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for Active button 
        ACTIVE_frame.place(x=610,y=260)
        ACTIVE_frame.place_configure(width = 390, height = 70)

        MSG_CNTR_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for MSG_CNTR button 
        MSG_CNTR_frame.place(x=80,y=180)
        MSG_CNTR_frame.place_configure(width = 390, height = 70)

        FAILSAFE_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for FAILSAFE button 
        FAILSAFE_frame.place(x=80,y=100)
        FAILSAFE_frame.place_configure(width = 390, height = 70)

        GATEWAY_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for GATEWAY button 
        GATEWAY_frame.place(x=610,y=180)
        GATEWAY_frame.place_configure(width = 390, height = 70)

        BUSOFF_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for BUSOFF button 
        BUSOFF_frame.place(x=610 ,y=20)
        BUSOFF_frame.place_configure(width = 390, height = 70)
                         
        CONFIG_CHECK_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for BUSOFF button 
        CONFIG_CHECK_frame.place(x=80 ,y=260)
        CONFIG_CHECK_frame.place_configure(width = 390, height = 70)
        
##        Maneuver_Testing_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
##                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for BUSOFF button 
##        Maneuver_Testing_frame.place(x=610 ,y=100)
##        Maneuver_Testing_frame.place_configure(width = 390, height = 70)
        
        ICC_cancel_Testing_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for BUSOFF button 
        ICC_cancel_Testing_frame.place(x=610 ,y=100)
        ICC_cancel_Testing_frame.place_configure(width = 390, height = 70)
        
        OK_frame= Frame(frame3,self,relief = RAISED, borderwidth = 2,
                           bg = browse_frame_color, height = 20)                                                              # Create a new frame for BUSOFF button 
        OK_frame.place(x=420 ,y=420)
        OK_frame.place_configure(width = 150, height = 70)          
################################################################################################################################
        
#######################################Checknbutton Creation #####################################################################
        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)
        ALL_check_button = Tkinter.Checkbutton(ALL_frame, text = "ALL",
                                               bg=frame1_color,
                                               width = 60,height = 3,
                                               relief = GROOVE,activebackground = "green",
                                               indicatoron=0,selectcolor=start_button_color,
                                               cursor = "hand2",font = "times 20 bold",
                                               onvalue = 1,offvalue = 0,variable=v1,command=Assign)
        
        ALL_check_button.pack(side=RIGHT,anchor=NE,padx=10,pady=10 )
        
        ALL_logo = Image.open(Image_Path+"All.JPG")
        ALL_logo = ALL_logo.resize((50, 40), Image.ANTIALIAS)
        ALL_image = ImageTk.PhotoImage(ALL_logo)
        ALL_label = Label(ALL_frame, image = ALL_image, bg = heading_color)
        ALL_label.image = ALL_image
        ALL_label.place(x = 8, y = 11)

        
        
        ITS_check_button = Tkinter.Checkbutton(ITS_frame, text = "ITS Application",                                #Creating ITS button
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v2,command=Assign)
        ITS_check_button.pack(side=RIGHT,anchor= NE,padx=10,pady=10)
        
        ITS_logo = Image.open(Image_Path+"ITS.JPG")
        ITS_logo = ITS_logo.resize((50, 40), Image.ANTIALIAS)
        ITS_image = ImageTk.PhotoImage(ITS_logo)
        ITS_label = Label(ITS_frame, image = ITS_image, bg = heading_color)
        ITS_label.image = ITS_image
        ITS_label.place(x = 8, y = 11)
        
        ACTIVE_check_button = Tkinter.Checkbutton(ACTIVE_frame, text = "Active Testing",                                    #Creating Active check button
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v3,command=Assign)
        ACTIVE_check_button.pack(side=RIGHT,anchor=NE,padx=10,pady=10)

        ACTIVE_logo = Image.open(Image_Path+"ddt2000.jpg")
        ACTIVE_logo = ACTIVE_logo.resize((50, 40), Image.ANTIALIAS)
        ACTIVE_image = ImageTk.PhotoImage(ACTIVE_logo)
        ACTIVE_label = Label(ACTIVE_frame, image = ACTIVE_image, bg = heading_color)
        ACTIVE_label.image = ACTIVE_image
        ACTIVE_label.place(x = 8, y = 11)
        
        MSG_COUNTER_check_button = Tkinter.Checkbutton(MSG_CNTR_frame, text = "Message Counter",
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v4,command=Assign)
        MSG_COUNTER_check_button.pack(side=RIGHT,anchor=NE ,padx=10,pady=10)

        MSG_CNTR_logo = Image.open(Image_Path+"MSG_CNTR.JPG")
        MSG_CNTR_logo = MSG_CNTR_logo.resize((50, 40), Image.ANTIALIAS)
        MSG_CNTR_image = ImageTk.PhotoImage(MSG_CNTR_logo)
        MSG_CNTR_label = Label(MSG_CNTR_frame, image = MSG_CNTR_image, bg = heading_color)
        MSG_CNTR_label.image = MSG_CNTR_image
        MSG_CNTR_label.place(x = 8, y = 11)
        
        FAILSAFE_check_button = Tkinter.Checkbutton( FAILSAFE_frame, text = "Failsafe Testing",
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v5,command=Assign)
        FAILSAFE_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10)

        FAILSAFE_logo = Image.open(Image_Path+"FAILSAFE.JPG")
        FAILSAFE_logo = FAILSAFE_logo.resize((50, 40), Image.ANTIALIAS)
        FAILSAFE_image = ImageTk.PhotoImage(FAILSAFE_logo)
        FAILSAFE_label = Label(FAILSAFE_frame, image = FAILSAFE_image, bg = heading_color)
        FAILSAFE_label.image = FAILSAFE_image
        FAILSAFE_label.place(x = 8, y = 11)
        
        GATEWAY_check_button = Tkinter.Checkbutton(GATEWAY_frame, text = "Gateway",
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v6,command=Assign)
        GATEWAY_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10)

        GATEWAY_logo = Image.open(Image_Path+"Gateway.jpg")
        GATEWAY_logo = GATEWAY_logo.resize((50, 40), Image.ANTIALIAS)
        GATEWAY_image = ImageTk.PhotoImage(GATEWAY_logo)
        GATEWAY_label = Label(GATEWAY_frame, image = GATEWAY_image, bg = heading_color)
        GATEWAY_label.image = GATEWAY_image
        GATEWAY_label.place(x = 8, y = 11)
        
        BUSOFF_check_button = Tkinter.Checkbutton(BUSOFF_frame, text = "Bus Off",
                                         bg=frame1_color,
                                                 
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v7,command=Assign)
        BUSOFF_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10)

        BUSOFF_logo = Image.open(Image_Path+"BusOff.jpg")
        BUSOFF_logo = BUSOFF_logo.resize((50, 40), Image.ANTIALIAS)
        BUSOFF_image = ImageTk.PhotoImage(BUSOFF_logo)
        BUSOFF_label = Label(BUSOFF_frame, image = BUSOFF_image, bg = heading_color)
        BUSOFF_label.image = BUSOFF_image
        BUSOFF_label.place(x = 8, y = 11)
        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)

        CONFIG_CHECK_check_button = Tkinter.Checkbutton(CONFIG_CHECK_frame, text = "Config Check",
                                         bg=frame1_color,
                                                 
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v8,command=Assign)
        CONFIG_CHECK_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10)

        CONFIG_CHECK_logo = Image.open(Image_Path+"ConfigCheck.jpg")
        CONFIG_CHECK_logo = CONFIG_CHECK_logo.resize((50, 40), Image.ANTIALIAS)
        CONFIG_CHECK_image = ImageTk.PhotoImage(CONFIG_CHECK_logo)
        CONFIG_CHECK_label = Label(CONFIG_CHECK_frame, image = CONFIG_CHECK_image, bg = heading_color)
        CONFIG_CHECK_label.image = CONFIG_CHECK_image
        CONFIG_CHECK_label.place(x = 8, y = 11)
        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)
##
##
##
##        Maneuver_Testing_check_button = Tkinter.Checkbutton(Maneuver_Testing_frame, text = "Maneuver Testing",
##                                         bg=frame1_color,
##                                                 
##                                     width = 60,
##                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
##                                         
##                                     activebackground = "green", cursor = "hand2",
##
##                                     font = "times 20 bold",
##                                           onvalue = 1,offvalue = 0,variable=v9,command=Assign)
##        Maneuver_Testing_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10) 
##
##
##        Maneuver_Testing_logo = Image.open(Image_Path+"BSW.jpg")
##        Maneuver_Testing_logo = Maneuver_Testing_logo.resize((50, 40), Image.ANTIALIAS)
##        Maneuver_Testing_image = ImageTk.PhotoImage(Maneuver_Testing_logo)
##        Maneuver_Testing_label = Label(Maneuver_Testing_frame, image = Maneuver_Testing_image, bg = heading_color)
##        Maneuver_Testing_label.image = Maneuver_Testing_image
##        Maneuver_Testing_label.place(x = 8, y = 11)
##        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)
       
        ICC_Cancel_Testing_button = Tkinter.Checkbutton(ICC_cancel_Testing_frame, text = "ICC Cancel Testing ",
                                         bg=frame1_color,
                                                 
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v9,command=Assign)
        ICC_Cancel_Testing_button.pack(side=LEFT,anchor=NE,padx=10,pady=10) 


        ICC_cancel_Testing_logo = Image.open(Image_Path+"BSW.jpg")
        ICC_cancel_Testing_logo = ICC_cancel_Testing_logo.resize((50, 40), Image.ANTIALIAS)
        ICC_cancel_Testing_image = ImageTk.PhotoImage(ICC_cancel_Testing_logo)
        ICC_cancel_Testing_label = Label(ICC_cancel_Testing_frame, image = ICC_cancel_Testing_image, bg = heading_color)
        ICC_cancel_Testing_label.image = ICC_cancel_Testing_image
        ICC_cancel_Testing_label.place(x = 8, y = 11)
        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)

        OK_button = Tkinter.Button(OK_frame, text = "OK",                                    #Creating Active check button
                                         bg=Browse_button_color,
                                     width = 15,
                                     height = 2, relief = RAISED,
                                      borderwidth=4,   
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",command=OK_Pressed )

        #OK_button.place(x=420 ,y=360)
        OK_button.pack(padx=5,pady=5)




        Plantmodel_button["state"] = DISABLED
        dispatch_button["state"] = DISABLED
        dispatch_button["state"] = NORMAL
        start_button["state"] = NORMAL
        stop_button["state"] = DISABLED
        reset_button["state"] = DISABLED
        
        ALL_check_button["state"] = NORMAL
        ITS_check_button["state"] = NORMAL
        FAILSAFE_check_button["state"] = NORMAL
        GATEWAY_check_button["state"] = NORMAL
        BUSOFF_check_button["state"] = NORMAL
        ACTIVE_check_button["state"] =  NORMAL
        MSG_COUNTER_check_button["state"] = NORMAL
        CONFIG_CHECK_check_button["state"] = NORMAL
       ## Maneuver_Testing_check_button["state"] = NORMAL
        ICC_Cancel_Testing_button["state"] = NORMAL
        OK_button["state"] = NORMAL
        


        
class ITS_Testing(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_1()

#** Definition of initializeUI (GUI creation) **#   

    def initializeUI_1(self):
        
        global tree_frame, tree, Script_Path, Org_Path
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()

    
    

#***********************************************#       
    
        global frame3
        global frame6

        global ITS_test_output_frame,ITS_vehicle_id_label,ITS_variant_label,ITS_application_label,ITS_testcase_label,ITS_result_label,ITS_Progress_label,ITS_vehicle_id_entry,ITS_variant_entry,ITS_application_entry,ITS_testcase_entry,ITS_result_entry
        global ITS_progressbar_color,ITS_progressbar_style,ITS_progressbar,ITS_overall_progressbar
        global heading_color,Browse_button_color
                           
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        background_frame.pack(fill=BOTH, expand=True)

        heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        heading_frame = Frame(background_frame,relief=RAISED,borderwidth=3,
                              bg = heading_color)                                                                       # Creating and Setting the design of Heading Frame
        heading_frame.place(x=15, y=15)
        heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(heading_frame, image = Nissan_image, bg = heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(heading_frame, image = TCS_image, bg = heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        heading_label= Label(heading_frame, justify="center",
                             text = "ITS APPLICATION TESTING",                                                 # Creating the Heading Label
                             bg =heading_color, fg = "black", font="Times 25 bold")
        heading_label.pack()        
        
       
        frame3 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =browse_frame_color,height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame3.pack(fill = BOTH, expand = TRUE)
##    
       
        ITS_test_output_frame = LabelFrame(frame3, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =430, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        ITS_test_output_frame.place(x = 505, y = 5)
        ITS_test_output_frame.grid_propagate(0)

        ITS_test_output_frame.columnconfigure(0, pad = 10)
        ITS_test_output_frame.columnconfigure(1, pad = 10)

        ITS_test_output_frame.rowconfigure(0, pad = 9)
        ITS_test_output_frame.rowconfigure(1, pad = 9)
        ITS_test_output_frame.rowconfigure(2, pad = 9)
        ITS_test_output_frame.rowconfigure(3, pad = 9)
        ITS_test_output_frame.rowconfigure(4, pad = 9)
        ITS_test_output_frame.rowconfigure(5, pad = 9)
        
        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        ITS_vehicle_id_label = Label(ITS_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        ITS_vehicle_id_label.grid(row = 0, column = 0)
        ITS_variant_label = Label(ITS_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ITS_variant_label.grid(row = 1, column = 0)
        ITS_application_label = Label (ITS_test_output_frame, text = "Application",
                                   width = 11, height = 2, bg = logging_labels_color,
                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ITS_application_label.grid(row = 2, column = 0)
        ITS_testcase_label = Label (ITS_test_output_frame, text = "Test Case No.",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ITS_testcase_label.grid(row = 3, column = 0)
        ITS_result_label = Label (ITS_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ITS_result_label.grid(row = 4, column = 0)
        ITS_Progress_label = Label (ITS_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        ITS_Progress_label.grid(row = 5, column = 0)
        
        ITS_vehicle_id_entry = Entry(ITS_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        ITS_vehicle_id_entry.grid(row = 0, column = 1)
        ITS_variant_entry = Entry(ITS_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        ITS_variant_entry.grid(row = 1, column = 1)
        ITS_application_entry = Entry(ITS_test_output_frame, width = 55, bd = 4,
                                  font = ("Arial", 10, "bold"))
        ITS_application_entry.grid(row = 2, column = 1)
        ITS_testcase_entry = Entry(ITS_test_output_frame, width = 55, bd = 4,
                               font = ("Arial", 10, "bold"))
        ITS_testcase_entry.grid(row = 3, column = 1)
        ITS_result_entry = Entry(ITS_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        ITS_result_entry.grid(row = 4, column = 1)

        ITS_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        ITS_progressbar_style = ttk.Style()
        ITS_progressbar_style.theme_use("default")
        ITS_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = ITS_progressbar_color)
        ITS_progressbar = ttk.Progressbar(ITS_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        ITS_progressbar.grid(row = 5, column = 1)

        
        ITS_logging_frame = LabelFrame(frame3, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =130, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame
        ITS_logging_frame.place(x = 505, y = 365)
        ITS_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(ITS_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 65)

        ITS_overall_progressbar = ttk.Progressbar(ITS_logging_frame,
                                              orient ="horizontal",
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        ITS_overall_progressbar.place (x = 120, y = 30)
        
        testinprogress_entry = Entry(ITS_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)                     
        
        tree_frame = Frame(frame3, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame.place(x = 15, y = 5)
        tree_frame.place_configure(width = 450, height = 490)

        scrollbar = Scrollbar(tree_frame, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        Data = ''
        tree = ttk.Treeview(tree_frame, height=25)                                                  # Tree data
        tree.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree.yview)
        tree.config(yscrollcommand=scrollbar.set)
        tree.pack(padx = 3,pady = 3)
            
        ITS_logo = Image.open(Image_Path+"ITS.JPG")                                                                          # Importing the dspace bench setup image
        ITS_logo=ITS_logo.resize((445, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(ITS_logo)
        label3 = Label(self, image = logo_pi, relief = RAISED)
        label3.image = logo_pi
        label3.place(x = 20, y = 70)

        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)        
        



class Active_Testing(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_2()

                                                                      # Thread Creation
    
    def initializeUI_2(self):
        
        global tree_frame, tree, Script_Path, Org_Path
        global ACTIVE_test_output_frame,ACTIVE_vehicle_id_label,ACTIVE_variant_label,test_case_type_label,testcase_ID_label,ACTIVE_result_label,ACTIVE_Progress_label,ACTIVE_vehicle_id_entry,ACTIVE_variant_entry,ACTIVE_application_entry,ACTIVE_testcase_entry,ACTIVE_result_entry
        global ACTIVE_progressbar_color,ACTIVE_progressbar_style,ACTIVE_progressbar
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()
        
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        background_frame.pack(fill=BOTH, expand=True)

        heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        heading_frame = Frame(background_frame,relief=RAISED,borderwidth=3,
                              bg = heading_color)                                                                       # Creating and Setting the design of Heading Frame
        heading_frame.place(x=15, y=15)
        heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(heading_frame, image = Nissan_image, bg = heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)

        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(heading_frame, image = TCS_image, bg = heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        heading_label= Label(heading_frame, justify="center",
                             text = "Active Testing",                                                 # Creating the Heading Label
                             bg =heading_color, fg = "black", font="Times 25 bold")
        heading_label.pack()  
      
        global frame5, result_entry_active,Browse_button_color
      
        frame5 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame5.pack(fill = BOTH, expand = TRUE)
          
        ACTIVE_test_output_frame = LabelFrame(frame5, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =400, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        ACTIVE_test_output_frame.place(x = 505, y = 5)
        ACTIVE_test_output_frame.grid_propagate(0)

        ACTIVE_test_output_frame.columnconfigure(0, pad = 10)
        ACTIVE_test_output_frame.columnconfigure(1, pad = 10)

        ACTIVE_test_output_frame.rowconfigure(0, pad = 9)
        ACTIVE_test_output_frame.rowconfigure(1, pad = 9)
        ACTIVE_test_output_frame.rowconfigure(2, pad = 9)
        ACTIVE_test_output_frame.rowconfigure(3, pad = 9)
        ACTIVE_test_output_frame.rowconfigure(4, pad = 9)
        ACTIVE_test_output_frame.rowconfigure(5, pad = 9)
        
        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        ACTIVE_vehicle_id_label = Label(ACTIVE_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        ACTIVE_vehicle_id_label.grid(row = 0, column = 0)
        ACTIVE_variant_label = Label(ACTIVE_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ACTIVE_variant_label.grid(row = 1, column = 0)
        ACTIVE_ECU_SW_Ver_label = Label (ACTIVE_test_output_frame, text = "ECU Software \n Version",
                                   width = 11, height = 2, bg = logging_labels_color,
                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ACTIVE_ECU_SW_Ver_label.grid(row = 2, column = 0)
        ACTIVE_testcase_ID_label = Label (ACTIVE_test_output_frame, text = "Test Case No.",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ACTIVE_testcase_ID_label.grid(row = 3, column = 0)
        ACTIVE_result_label = Label (ACTIVE_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ACTIVE_result_label.grid(row = 4, column = 0)
        ACTIVE_Progress_label = Label (ACTIVE_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        ACTIVE_Progress_label.grid(row = 5, column = 0)
        
        ACTIVE_vehicle_id_entry = Entry(ACTIVE_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        ACTIVE_vehicle_id_entry.grid(row = 0, column = 1)
        ACTIVE_variant_entry = Entry(ACTIVE_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        ACTIVE_variant_entry.grid(row = 1, column = 1)
        ACTIVE_ECU_SW_Ver_entry = Entry(ACTIVE_test_output_frame, width = 55, bd = 4,
                                  font = ("Arial", 10, "bold"))
        ACTIVE_ECU_SW_Ver_entry.grid(row = 2, column = 1)
        ACTIVE_testcase_entry = Entry(ACTIVE_test_output_frame, width = 55, bd = 4,
                               font = ("Arial", 10, "bold"))
        ACTIVE_testcase_entry.grid(row = 3, column = 1)
        ACTIVE_result_entry = Entry(ACTIVE_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        ACTIVE_result_entry.grid(row = 4, column = 1)

        ACTIVE_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        ACTIVE_progressbar_style = ttk.Style()
        ACTIVE_progressbar_style.theme_use("default")
        ACTIVE_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = ACTIVE_progressbar_color)
        ACTIVE_progressbar = ttk.Progressbar(ACTIVE_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        ACTIVE_progressbar.grid(row = 5, column = 1)
        
        ACTIVE_logging_frame = LabelFrame(frame5, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =145, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame for Active Testing 
        ACTIVE_logging_frame.place(x = 505, y = 355)
        ACTIVE_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(ACTIVE_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 73)

        ACTIVE_overall_progressbar = ttk.Progressbar(ACTIVE_logging_frame,
                                              orient ="horizontal",                                                   ## Creating and Setting Overall progress bar for Active Testing 
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        ACTIVE_overall_progressbar.place (x = 120, y = 40)
        
        testinprogress_entry = Entry(ACTIVE_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)                     
        
        tree_frame_1 = Frame(frame5, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame_1.place(x = 15, y = 5)
        tree_frame_1.place_configure(width = 450, height = 490)

        scrollbar = Scrollbar(tree_frame_1, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        Data = ''
        tree1 = ttk.Treeview(tree_frame_1, height=25)                                                  # Tree data
        tree1.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree1.yview)
        tree1.config(yscrollcommand=scrollbar.set)
        tree1.pack(padx = 3,pady = 3)

        Active_logo = Image.open(Image_Path+"ddt2000.jpg")                                                                          # Importing the dspace bench setup image
        Active_logo=Active_logo.resize((440, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(Active_logo)
        label2 = Label(self, image = logo_pi, relief = RAISED)
        label2.image = logo_pi
        label2.place(x = 20, y = 70)         
            
        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)     
            
       
        

class FAILSAFE(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_3()

                                                                      # Thread Creation
    
    def initializeUI_3(self):
        
        global tree_frame, tree, Script_Path, Org_Path, ECU_sensor_drop, failsafe_cat__drop
        global FAILSAFE_test_output_frame,FAILSAFE_vehicle_id_label,FAILSAFE_variant_label,ECU_label,CAN_ID_label,FAILSAFE_CAT_label,FAILSAFE_result_label,FAILSAFE_Progress_label,FAILSAFE_vehicle_id_entry,FAILSAFE_variant_entry,FAILSAFE_ECU_entry,FAILSAFE_CAN_ID_entry,FAILSAFE_CAT_entry,FAILSAFE_result_entry,FAILSAFE_overall_progressbar
        global FAILSAFE_progressbar_color,FAILSAFE_progressbar_style,FAILSAFE_progressbar,background_color
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()
        
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        background_frame.pack(fill=BOTH, expand=True)

        heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        heading_frame = Frame(background_frame,relief=RAISED,borderwidth=3,
                              bg = heading_color)                                                                       # Creating and Setting the design of Heading Frame
        heading_frame.place(x=15, y=15)
        heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(heading_frame, image = Nissan_image, bg = heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(heading_frame, image = TCS_image, bg = heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        heading_label= Label(heading_frame, justify="center",
                             text = "Failsafe Testing",                                                 # Creating the Heading Label
                             bg =heading_color, fg = "black", font="Times 25 bold")
        heading_label.pack()

        dropdown_frame = Frame (background_frame, relief = RAISED,
                               borderwidth = 2, bg = background_color)                                              # Creating a frame for dropdown Menu 
        dropdown_frame.place(x = 15, y = 100)
        dropdown_frame.place_configure(width = 450, height = 100)
        
        ECU_sensor_drop_label=Label(dropdown_frame, justify="left",
                             text = "ECU/sensor",                                                                  #Creating the  Label for ECU/sensor label
                             bg =background_color, fg = "black", font="Times 16 bold")
        ECU_sensor_drop_label.place(x = 50, y = 5)

        failsafe_cat__drop_label=Label(dropdown_frame, justify="left",
                             text = "Failsafe Category",                                                            # Creating the Failsafe category Label
                             bg =background_color, fg = "black", font="Times 16 bold")
        failsafe_cat__drop_label.place(x = 250, y = 5)        
        
        ECU_sensor_drop_frame=Frame (dropdown_frame, relief = RAISED,
                               borderwidth = 2, bg = background_color)                                              # Creating a frame for ECU/sensor dropdown option
        ECU_sensor_drop_frame.place(x = 5, y = 40)
        ECU_sensor_drop_frame.configure(width=10,height=2)
                
      
        failsafe_cat__drop_frame=Frame (dropdown_frame, relief = RAISED,
                               borderwidth = 2, bg = background_color)                                              # Creating a frame for failsafe Category option
        failsafe_cat__drop_frame.place(x = 230, y = 40)
        failsafe_cat__drop_frame.configure(width=10,height=2)
        
        global frame7
        
        frame7 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame7.pack(fill = BOTH, expand = TRUE)
     
        
        FAILSAFE_test_output_frame = LabelFrame(frame7, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =370, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        FAILSAFE_test_output_frame.place(x = 505, y = 5)
        FAILSAFE_test_output_frame.grid_propagate(0)

        FAILSAFE_test_output_frame.columnconfigure(0, pad = 10)
        FAILSAFE_test_output_frame.columnconfigure(1, pad = 10)

        FAILSAFE_test_output_frame.rowconfigure(0, pad = 9)
        FAILSAFE_test_output_frame.rowconfigure(1, pad = 9)
        FAILSAFE_test_output_frame.rowconfigure(2, pad = 9)
        FAILSAFE_test_output_frame.rowconfigure(3, pad = 9)
        FAILSAFE_test_output_frame.rowconfigure(4, pad = 9)
        FAILSAFE_test_output_frame.rowconfigure(5, pad = 9)
        FAILSAFE_test_output_frame.rowconfigure(6, pad = 9)

        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        FAILSAFE_vehicle_id_label = Label(FAILSAFE_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        FAILSAFE_vehicle_id_label.grid(row = 0, column = 0)
        FAILSAFE_variant_label = Label(FAILSAFE_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        FAILSAFE_variant_label.grid(row = 1, column = 0)
        ECU_label = Label (FAILSAFE_test_output_frame, text = "ECU/Sensor",
                                   width = 11, height = 2, bg = logging_labels_color,
                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ECU_label.grid(row = 3, column = 0)
        CAN_ID_label = Label (FAILSAFE_test_output_frame, text = "CAN ID",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        CAN_ID_label.grid(row = 4, column = 0)
        
        FAILSAFE_CAT_label = Label (FAILSAFE_test_output_frame, text = "Failsafe \n category ",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        FAILSAFE_CAT_label.grid(row = 2, column = 0)
        
        FAILSAFE_result_label = Label (FAILSAFE_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        FAILSAFE_result_label.grid(row = 5, column = 0)
        FAILSAFE_Progress_label = Label (FAILSAFE_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        FAILSAFE_Progress_label.grid(row = 6, column = 0)
        
        FAILSAFE_vehicle_id_entry = Entry(FAILSAFE_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        FAILSAFE_vehicle_id_entry.grid(row = 0, column = 1)
        FAILSAFE_variant_entry = Entry(FAILSAFE_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        FAILSAFE_variant_entry.grid(row = 1, column = 1)
        FAILSAFE_ECU_entry = Entry(FAILSAFE_test_output_frame, width = 55, bd = 4,
                                  font = ("Arial", 10, "bold"))
        FAILSAFE_ECU_entry.grid(row = 3, column = 1)
        FAILSAFE_CAN_ID_entry = Entry(FAILSAFE_test_output_frame, width = 55, bd = 4,
                               font = ("Arial", 10, "bold"))
        FAILSAFE_CAN_ID_entry.grid(row = 4, column = 1)
        FAILSAFE_CAT_entry = Entry(FAILSAFE_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        FAILSAFE_CAT_entry.grid(row = 2, column = 1)
        
        FAILSAFE_result_entry = Entry(FAILSAFE_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        FAILSAFE_result_entry.grid(row = 5, column = 1)

        FAILSAFE_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        FAILSAFE_progressbar_style = ttk.Style()
        FAILSAFE_progressbar_style.theme_use("default")
        FAILSAFE_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = FAILSAFE_progressbar_color)
        FAILSAFE_progressbar = ttk.Progressbar(FAILSAFE_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        FAILSAFE_progressbar.grid(row = 6, column = 1)

        FAILSAFE_logging_frame = LabelFrame(frame7, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =145, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame
        FAILSAFE_logging_frame.place(x = 505, y = 355)
        FAILSAFE_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(FAILSAFE_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 73)

        FAILSAFE_overall_progressbar = ttk.Progressbar(FAILSAFE_logging_frame,
                                              orient ="horizontal",                                                     #Failsafe Overall progressbar Creation
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        FAILSAFE_overall_progressbar.place (x = 120, y = 40)
        
        testinprogress_entry = Entry(FAILSAFE_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)           
        
        tree_frame_2 = Frame(frame7, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame_2.place(x = 15, y = 5)
        tree_frame_2.place_configure(width = 450, height = 490)

        #option_frame = Frame(background_frame,relief = RAISED, borderwidth=2,
                          # bg =background_color )                                                                     # Creating and Setting  tree frame 
        #option_frame.place(x = 15, y = 80)
        #option_frame.place_configure(width = 450, height = 180)
        

        scrollbar = Scrollbar(tree_frame_2, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        Data = ''
        tree2 = ttk.Treeview(tree_frame_2, height=25)                                                  # Tree data
        tree2.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree2.yview)
        tree2.config(yscrollcommand=scrollbar.set)
        tree2.pack(padx = 3,pady = 3)
        
        logo = Image.open(Image_Path+"img_to_import.JPG")                                                                                # Importing the Failsafe logo image 
        logo=logo.resize((550, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 506, y = 70)

        FAILSAFE_logo = Image.open(Image_Path+"FAILSAFE.JPG")                                                                          # Importing the dspace bench setup image
        FAILSAFE_logo=FAILSAFE_logo.resize((445, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(FAILSAFE_logo)
        label3 = Label(self, image = logo_pi, relief = RAISED)
        label3.image = logo_pi
        label3.place(x = 18, y = 70)
      
##############################Edited################################


class Message_counter(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_Msg()

                                                                    
    
    def initializeUI_Msg(self):
        
        global tree_frame, tree, Script_Path, Org_Path,Message_Counter_Tree,tree_frame_3
        global Message_counter_vehicle_id_entry,Message_counter_variant_entry,Message_counter_CAN_ID_entry,Message_counter_SampleTime_entry,Message_counter_result_entry,Message_counter_CANchannel_entry,Message_counter_progressbar,Message_counter_overall_progressbar

        global Message_counter_test_output_frame,Message_counter_vehicle_id_label,Message_counter_vehicle_id_label,Message_counter_variant_label,Message_counter_CANchannel_label,Message_counter_CAN_ID_label,Message_counter_SampleTime_label,Message_counter_result_label,Message_counter_Progress_label
        global Message_counter_background_frame
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()
        
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        Message_counter_background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        Message_counter_background_frame.pack(fill=BOTH, expand=True)

        Message_counter_heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        Message_counter_heading_frame = Frame(Message_counter_background_frame,relief=RAISED,borderwidth=3,
                              bg = Message_counter_heading_color)                                                                       # Creating and Setting the design of Heading Frame
        Message_counter_heading_frame.place(x=15, y=15)
        Message_counter_heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(Message_counter_heading_frame, image = Nissan_image, bg = Message_counter_heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(Message_counter_heading_frame, image = TCS_image, bg = Message_counter_heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        Message_counter_heading_label= Label(Message_counter_heading_frame, justify="center",
                             text = "Message Counter",                                                 # Creating the Heading Label
                             bg =Message_counter_heading_color, fg = "black", font="Times 25 bold")
        Message_counter_heading_label.pack()
        
##        buttons_frame = Frame (background_frame, relief = RAISED,
##                               borderwidth = 2, bg = frame1_color)                                                      # Creating a frame for start, stop and reset buttons
##        buttons_frame.place(x = 15, y = 265)
##        buttons_frame.place_configure(width = 450, height = 60)

   
        
        global frame8
        
        frame8 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame8.pack(fill = BOTH, expand = TRUE)

##        frame8 = Frame(self,relief = RAISED, borderwidth = 2,
##                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
##        frame8.pack(fill = BOTH, expand = TRUE)        
        
        Message_counter_test_output_frame = LabelFrame(frame8, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =370, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        Message_counter_test_output_frame.place(x = 505, y = 5)
        Message_counter_test_output_frame.grid_propagate(0)

        Message_counter_test_output_frame.columnconfigure(0, pad = 10)
        Message_counter_test_output_frame.columnconfigure(1, pad = 10)

        Message_counter_test_output_frame.rowconfigure(0, pad = 9)
        Message_counter_test_output_frame.rowconfigure(1, pad = 9)
        Message_counter_test_output_frame.rowconfigure(2, pad = 9)
        Message_counter_test_output_frame.rowconfigure(3, pad = 9)
        Message_counter_test_output_frame.rowconfigure(4, pad = 9)
        Message_counter_test_output_frame.rowconfigure(5, pad = 9)
        Message_counter_test_output_frame.rowconfigure(6, pad = 9)

        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        Message_counter_vehicle_id_label = Label(Message_counter_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        Message_counter_vehicle_id_label.grid(row = 0, column = 0)
        
        Message_counter_variant_label = Label(Message_counter_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        Message_counter_variant_label.grid(row = 1, column = 0)
        

        
        Message_counter_CAN_ID_label = Label (Message_counter_test_output_frame, text = "CAN ID",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        Message_counter_CAN_ID_label.grid(row = 2, column = 0)
        

        
        Message_counter_result_label = Label (Message_counter_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        Message_counter_result_label.grid(row = 3, column = 0)
        
        Message_counter_Progress_label = Label (Message_counter_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        Message_counter_Progress_label.grid(row = 4, column = 0)
        
        Message_counter_vehicle_id_entry = Entry(Message_counter_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        Message_counter_vehicle_id_entry.grid(row = 0, column = 1)
        Message_counter_variant_entry = Entry(Message_counter_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        Message_counter_variant_entry.grid(row = 1, column = 1)

        Message_counter_CAN_ID_entry = Entry(Message_counter_test_output_frame, width = 55, bd = 4,
                               font = ("Arial", 10, "bold"))
        Message_counter_CAN_ID_entry.grid(row = 2, column = 1)

        
        Message_counter_result_entry = Entry(Message_counter_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        Message_counter_result_entry.grid(row = 3, column = 1)

        Message_counter_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        Message_counter_progressbar_style = ttk.Style()
        Message_counter_progressbar_style.theme_use("default")
        Message_counter_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = FAILSAFE_progressbar_color)
        Message_counter_progressbar = ttk.Progressbar(Message_counter_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        Message_counter_progressbar.grid(row = 4, column = 1)

        Message_counter_logging_frame = LabelFrame(frame8, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =140, relief = RAISED, font = "Helvetica")                                            # Creating and Setting Overall progress frame
        Message_counter_logging_frame.place(x = 505, y = 355)
        Message_counter_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(Message_counter_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 70)

        Message_counter_overall_progressbar = ttk.Progressbar(Message_counter_logging_frame,
                                              orient ="horizontal",
                                              mode = "determinate",                                                             #Creating message counter overall progressbar
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        Message_counter_overall_progressbar.place (x = 120, y = 40)

        testinprogress_entry = Entry(Message_counter_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)           
        
        tree_frame_3 = Frame(frame8, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame_3.place(x = 15, y = 5)
        tree_frame_3.place_configure(width = 450, height = 490)



        scrollbar = Scrollbar(tree_frame_3, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        ## Data = ''
        Message_Counter_Tree = ttk.Treeview(tree_frame_3, height=25)                                                  # Tree data
        Message_Counter_Tree.column("#0",minwidth=0,width=450, stretch=NO)
        scrollbar.config(command=Message_Counter_Tree.yview)
        Message_Counter_Tree.config(yscrollcommand=scrollbar.set)
        Message_Counter_Tree.pack(padx = 3,pady = 3)

        
        logo = Image.open(Image_Path+"MSG_CNTR.JPG")                                                                                       # Importing the message counter logo image
        logo=logo.resize((440, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 20, y = 70)
        
        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                                  # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)
        return Message_counter_background_frame
class Gateway(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_Gateway()

                                                                    
    
    def initializeUI_Gateway(self):
        
        global tree_frame, tree, Script_Path, Org_Path
        
        global Gateway_test_output_frame,Gateway_vehicle_id_label,Gateway_variant_label,ECU_label,CAN_ID_label,Gateway_CAT_label,Gateway_result_label,Gateway_Progress_label,ECU_entry,Gateway_CAT_entry
        global Gateway_progressbar_color,Gateway_progressbar_style,Gateway_progressbar,Gateway_overall_progressbar
        global Gateway_vehicle_id_entry,Gateway_variant_entry,Gateway_CAN_ID_entry,Gateway_result_entry,Gateway_GWTGW_entry
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()
        
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        Gateway_background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        Gateway_background_frame.pack(fill=BOTH, expand=True)

        Gateway_heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        Gateway_heading_frame = Frame(Gateway_background_frame,relief=RAISED,borderwidth=3,
                              bg = Gateway_heading_color)                                                                       # Creating and Setting the design of Heading Frame
        Gateway_heading_frame.place(x=15, y=15)
        Gateway_heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(Gateway_heading_frame, image = Nissan_image, bg = Gateway_heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(Gateway_heading_frame, image = TCS_image, bg = Gateway_heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        Gateway_heading_label= Label(Gateway_heading_frame, justify="center",
                             text = "Gateway",                                                 # Creating the Heading Label
                             bg =Gateway_heading_color, fg = "black", font="Times 25 bold")
        Gateway_heading_label.pack()
        
##        buttons_frame = Frame (background_frame, relief = RAISED,
##                               borderwidth = 2, bg = frame1_color)                                                      # Creating a frame for start, stop and reset buttons
##        buttons_frame.place(x = 15, y = 265)
##        buttons_frame.place_configure(width = 450, height = 60)

   
        
        global frame9
        
        frame9 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame9.pack(fill = BOTH, expand = TRUE)

##        frame8 = Frame(self,relief = RAISED, borderwidth = 2,
##                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
##        frame8.pack(fill = BOTH, expand = TRUE)        
        
        Gateway_test_output_frame = LabelFrame(frame9, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =370, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        Gateway_test_output_frame.place(x = 505, y = 5)
        Gateway_test_output_frame.grid_propagate(0)

        Gateway_test_output_frame.columnconfigure(0, pad = 10)
        Gateway_test_output_frame.columnconfigure(1, pad = 10)

        Gateway_test_output_frame.rowconfigure(0, pad = 9)
        Gateway_test_output_frame.rowconfigure(1, pad = 9)
        Gateway_test_output_frame.rowconfigure(2, pad = 9)
        Gateway_test_output_frame.rowconfigure(3, pad = 9)
        Gateway_test_output_frame.rowconfigure(4, pad = 9)
        Gateway_test_output_frame.rowconfigure(5, pad = 9)
        Gateway_test_output_frame.rowconfigure(6, pad = 9)

        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        Gateway_vehicle_id_label = Label(Gateway_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        Gateway_vehicle_id_label.grid(row = 0, column = 0)
        Gateway_variant_label = Label(Gateway_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        Gateway_variant_label.grid(row = 1, column = 0)
##        Gateway_ECU_label = Label (Gateway_test_output_frame, text = "ECU/Sensor",
##                                   width = 11, height = 2, bg = logging_labels_color,
##                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
##        Gateway_ECU_label.grid(row = 2, column = 0)
        Gateway_CAN_ID_label = Label (Gateway_test_output_frame, text = "CAN ID",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        Gateway_CAN_ID_label.grid(row = 3, column = 0)
        
##        Gateway_CANchannel_label = Label (Gateway_test_output_frame, text = "CAN channel ",
##                                width =11, height = 2, bg = logging_labels_color,
##                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
##        Gateway_CANchannel_label.grid(row = 3, column = 0)
        
        Gateway_GWTGW_label = Label (Gateway_test_output_frame, text = "GW-TGW / \n GW-DIAG ",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        Gateway_GWTGW_label.grid(row = 2, column = 0)
        
        Gateway_result_label = Label (Gateway_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        Gateway_result_label.grid(row = 4, column = 0)
        Gateway_Progress_label = Label (Gateway_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        Gateway_Progress_label.grid(row = 5, column = 0)
        
        Gateway_vehicle_id_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        Gateway_vehicle_id_entry.grid(row = 0, column = 1)
        Gateway_variant_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        Gateway_variant_entry.grid(row = 1, column = 1)
        Gateway_CAN_ID_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
                                  font = ("Arial", 10, "bold"))
        Gateway_CAN_ID_entry.grid(row = 3, column = 1)
##        Gateway_CAN_ID_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
##                               font = ("Arial", 10, "bold"))
##        Gateway_CAN_ID_entry.grid(row = 3, column = 1)
##        Gateway_CANchannel_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
##                             font = ("Arial", 10, "bold"))
##        Gateway_CANchannel_entry.grid(row = 3, column = 1)
        Gateway_GWTGW_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        Gateway_GWTGW_entry.grid(row = 2, column = 1)
        
        Gateway_result_entry = Entry(Gateway_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        Gateway_result_entry.grid(row = 4, column = 1)

        Gateway_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        Gateway_progressbar_style = ttk.Style()
        Gateway_progressbar_style.theme_use("default")
        Gateway_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = Gateway_progressbar_color)
        Gateway_progressbar = ttk.Progressbar(Gateway_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        Gateway_progressbar.grid(row = 5, column = 1)

        Gateway_logging_frame = LabelFrame(frame9, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =145, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame
        Gateway_logging_frame.place(x = 505, y = 355)
        Gateway_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(Gateway_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 73)

        Gateway_overall_progressbar = ttk.Progressbar(Gateway_logging_frame,
                                              orient ="horizontal",
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        Gateway_overall_progressbar.place (x = 120, y = 40)
        
        testinprogress_entry = Entry(Gateway_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)                   
        
        tree_frame_4 = Frame(frame9, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame_4.place(x = 15, y = 5)
        tree_frame_4.place_configure(width = 450, height = 490)

##        option_frame = Frame(background_frame,relief = RAISED, borderwidth=2,
##                           bg =background_color )                                                                     # Creating and Setting  tree frame 
##        option_frame.place(x = 15, y = 80)
##        option_frame.place_configure(width = 450, height = 180)
        

        scrollbar = Scrollbar(tree_frame_4, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        Data = ''
        tree4 = ttk.Treeview(tree_frame_4, height=25)                                                  # Tree data
        tree4.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree4.yview)
        tree4.config(yscrollcommand=scrollbar.set)
        tree4.pack(padx = 3,pady = 3)

        logo = Image.open(Image_Path+"Gateway.jpg")                                                                        # Importing the dspace bench setup image
        logo=logo.resize((440, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 20, y = 70)
        
        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                  # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)
class BusOff(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent)

        self.parent = parent
        self.initializeUI_Busoff()



    def initializeUI_Busoff(self):

        global tree_frame_5, tree, Script_Path, Org_Path
        global Busoff_test_output_frame,Busoff_vehicle_id_label,Busoff_variant_label,Busoff_CAN_ID_label,Busoff_result_label,Busoff_Progress_label,Busoff_vehicle_id_entry,Busoff_variant_entry,Busoff_CAN_ID_entry,Busoff_result_entry
        global Busoff_progressbar_color,Busoff_progressbar_style,BusOff_progressbar,BusOff_overall_progressbar,Busoff_CANchannel_entry
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()

        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        Busoff_background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        Busoff_background_frame.pack(fill=BOTH, expand=True)

        Busoff_heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        Busoff_heading_frame = Frame(Busoff_background_frame,relief=RAISED,borderwidth=3,
                              bg = Busoff_heading_color)                                                                       # Creating and Setting the design of Heading Frame
        Busoff_heading_frame.place(x=15, y=15)
        Busoff_heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(Busoff_heading_frame, image = Nissan_image, bg = Busoff_heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(Busoff_heading_frame, image = TCS_image, bg = Busoff_heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)

        Busoff_heading_label= Label(Busoff_heading_frame, justify="center",
                             text = "Bus Off",                                                 # Creating the Heading Label
                             bg =Busoff_heading_color, fg = "black", font="Times 25 bold")
        Busoff_heading_label.pack()

##        buttons_frame = Frame (background_frame, relief = RAISED,
##                               borderwidth = 2, bg = frame1_color)                                                      # Creating a frame for start, stop and reset buttons
##        buttons_frame.place(x = 15, y = 265)
##        buttons_frame.place_configure(width = 450, height = 60)



        global frame10

        frame10 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window
        frame10.pack(fill = BOTH, expand = TRUE)


        Busoff_test_output_frame = LabelFrame(frame10, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =380, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        Busoff_test_output_frame.place(x = 505, y = 5)
        Busoff_test_output_frame.grid_propagate(0)

        Busoff_test_output_frame.columnconfigure(0, pad = 10)
        Busoff_test_output_frame.columnconfigure(1, pad = 10)

        Busoff_test_output_frame.rowconfigure(0, pad = 9)
        Busoff_test_output_frame.rowconfigure(1, pad = 9)
        Busoff_test_output_frame.rowconfigure(2, pad = 9)
        Busoff_test_output_frame.rowconfigure(3, pad = 9)
        Busoff_test_output_frame.rowconfigure(4, pad = 9)
        Busoff_test_output_frame.rowconfigure(5, pad = 9)
        Busoff_test_output_frame.rowconfigure(6, pad = 9)

        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        Busoff_test_output_frame_vehicle_id_label = Label(Busoff_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        Busoff_test_output_frame_vehicle_id_label.grid(row = 0, column = 0)
        Busoff_test_output_frame_variant_label = Label(Busoff_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        Busoff_test_output_frame_variant_label.grid(row = 1, column = 0)
        Busoff_CANchannel_label = Label (Busoff_test_output_frame, text = "CAN Channel",
                                   width = 11, height = 2, bg = logging_labels_color,
                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        Busoff_CANchannel_label.grid(row = 2, column = 0)
        Busoff_CAN_ID_label = Label (Busoff_test_output_frame, text = "CAN ID",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        Busoff_CAN_ID_label.grid(row = 3, column = 0)
##
##        Busoff_CAT_label = Label (Busoff_test_output_frame, text = "Failsafe \n category ",
##                                width =11, height = 2, bg = logging_labels_color,
##                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
##        Busoff_CAT_label.grid(row = 4, column = 0)

        Busoff_result_label = Label (Busoff_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        Busoff_result_label.grid(row = 4, column = 0)

        Busoff_Progress_label = Label (Busoff_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        Busoff_Progress_label.grid(row = 5, column = 0)

        Busoff_vehicle_id_entry = Entry(Busoff_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        Busoff_vehicle_id_entry.grid(row = 0, column = 1)
        Busoff_variant_entry = Entry(Busoff_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        Busoff_variant_entry.grid(row = 1, column = 1)
        Busoff_CANchannel_entry = Entry(Busoff_test_output_frame, width = 55, bd = 4,
                                  font = ("Arial", 10, "bold"))
        Busoff_CANchannel_entry.grid(row = 2, column = 1)
        Busoff_CAN_ID_entry = Entry(Busoff_test_output_frame, width = 55, bd = 4,
                               font = ("Arial", 10, "bold"))
        Busoff_CAN_ID_entry.grid(row = 3, column = 1)
##        Busoff_CAT_entry = Entry(Busoff_test_output_frame, width = 55, bd = 4,
##                             font = ("Arial", 10, "bold"))
##        Busoff_CAT_entry.grid(row = 4, column = 1)
##
        Busoff_result_entry = Entry(Busoff_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        Busoff_result_entry.grid(row = 4, column = 1)

        Busoff_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        Busoff_progressbar_style = ttk.Style()
        Busoff_progressbar_style.theme_use("default")
        Busoff_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = Busoff_progressbar_color)
        BusOff_progressbar = ttk.Progressbar(Busoff_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        BusOff_progressbar.grid(row = 5, column = 1)

        Busoff_logging_frame = LabelFrame(frame10, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =145, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame for Bus Off
        Busoff_logging_frame.place(x = 505, y = 355)
        Busoff_logging_frame.grid_propagate(FALSE)


        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(Busoff_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 73)

        BusOff_overall_progressbar = ttk.Progressbar(Busoff_logging_frame,
                                              orient ="horizontal",                                                      #Creating Overall progressbar for Bus-OFF
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)
        BusOff_overall_progressbar.place (x = 120, y = 40)

        testinprogress_entry = Entry(Busoff_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)

        tree_frame_5 = Frame(frame10, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame
        tree_frame_5.place(x = 15, y = 5)
        tree_frame_5.place_configure(width = 450, height = 490)

##        option_frame = Frame(background_frame,relief = RAISED, borderwidth=2,
##                           bg =background_color )                                                                     # Creating and Setting  tree frame
##        option_frame.place(x = 15, y = 80)
##        option_frame.place_configure(width = 450, height = 180)


        scrollbar = Scrollbar(tree_frame_5, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)


        Data = ''
        tree5 = ttk.Treeview(tree_frame_5, height=25)                                                  # Tree data
        tree5.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree5.yview)
        tree5.config(yscrollcommand=scrollbar.set)
        tree5.pack(padx = 3,pady = 3)


        logo = Image.open(Image_Path+"BusOff.jpg")                                                                                 # Importing the BusOff logo image
        logo=logo.resize((440, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 20, y = 70)

        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)        

class ConfigCheck(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_ConfigCheck()

                                                                    
    
    def initializeUI_ConfigCheck(self):
        
        global tree_frame, tree, Script_Path, Org_Path
        global ConfigCheck_test_output_frame,ConfigCheck_vehicle_id_label,ConfigCheck_variant_label,ECU_label,CAN_ID_label,ConfigCheck_CAT_label,ConfigCheck_result_label,FAILSAFE_Progress_label,FAILSAFE_vehicle_id_entry,FAILSAFE_variant_entry,ECU_entry,CAN_ID_entry,FAILSAFE_CAT_entry,FAILSAFE_result_entry
        global ConfigCheck_progressbar_color,ConfigCheck_progressbar_style,ConfigCheck_progressbar,ConfigCheck_vehicle_id_entry,ConfigCheck_variant,ConfigCheck_overall_progressbar,ConfigCheck_variant_entry,ConfigCheck_result_entry
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()
        
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        ConfigCheck_background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        ConfigCheck_background_frame.pack(fill=BOTH, expand=True)

        ConfigCheck_heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        ConfigCheck_heading_frame = Frame(ConfigCheck_background_frame,relief=RAISED,borderwidth=3,
                              bg = ConfigCheck_heading_color)                                                                       # Creating and Setting the design of Heading Frame
        ConfigCheck_heading_frame.place(x=15, y=15)
        ConfigCheck_heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(ConfigCheck_heading_frame, image = Nissan_image, bg = ConfigCheck_heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(ConfigCheck_heading_frame, image = TCS_image, bg = ConfigCheck_heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        ConfigCheck_heading_label= Label(ConfigCheck_heading_frame, justify="center",
                             text = "Config Check",                                                 # Creating the Heading Label
                             bg =ConfigCheck_heading_color, fg = "black", font="Times 25 bold")
        ConfigCheck_heading_label.pack()
        
##        buttons_frame = Frame (background_frame, relief = RAISED,
##                               borderwidth = 2, bg = frame1_color)                                                      # Creating a frame for start, stop and reset buttons
##        buttons_frame.place(x = 15, y = 265)
##        buttons_frame.place_configure(width = 450, height = 60)

   
        
        global frame11
        
        frame11 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =background_color, height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame11.pack(fill = BOTH, expand = TRUE)


        ConfigCheck_test_output_frame = LabelFrame(frame11, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =380, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        ConfigCheck_test_output_frame.place(x = 505, y = 5)
        ConfigCheck_test_output_frame.grid_propagate(0)

        ConfigCheck_test_output_frame.columnconfigure(0, pad = 10)
        ConfigCheck_test_output_frame.columnconfigure(1, pad = 10)

        ConfigCheck_test_output_frame.rowconfigure(0, pad = 9)
        ConfigCheck_test_output_frame.rowconfigure(1, pad = 9)
        ConfigCheck_test_output_frame.rowconfigure(2, pad = 9)
        ConfigCheck_test_output_frame.rowconfigure(3, pad = 9)
        ConfigCheck_test_output_frame.rowconfigure(4, pad = 9)
        ConfigCheck_test_output_frame.rowconfigure(5, pad = 9)
        ConfigCheck_test_output_frame.rowconfigure(6, pad = 9)

        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        ConfigCheck_test_output_frame_vehicle_id_label = Label(ConfigCheck_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        ConfigCheck_test_output_frame_vehicle_id_label.grid(row = 0, column = 0)
        ConfigCheck_test_output_frame_variant_label = Label(ConfigCheck_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ConfigCheck_test_output_frame_variant_label.grid(row = 1, column = 0)
##        ConfigCheck_CANchannel_label = Label (ConfigCheck_test_output_frame, text = "Torishi code",
##                                   width = 11, height = 2, bg = logging_labels_color,
##                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
##        ConfigCheck_CANchannel_label.grid(row = 2, column = 0)
#        ConfigCheck_CAN_ID_label = Label (ConfigCheck_test_output_frame, text = "CAN ID",
 #                               width =11, height = 2, bg = logging_labels_color,
 #                               relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
 #       ConfigCheck_CAN_ID_label.grid(row = 3, column = 0)
##        
##        ConfigCheck_CAT_label = Label (ConfigCheck_test_output_frame, text = "Failsafe \n category ",
##                                width =11, height = 2, bg = logging_labels_color,
##                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
##        ConfigCheck_CAT_label.grid(row = 4, column = 0)
        
        ConfigCheck_result_label = Label (ConfigCheck_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ConfigCheck_result_label.grid(row = 3, column = 0)
        
        ConfigCheck_Progress_label = Label (ConfigCheck_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        ConfigCheck_Progress_label.grid(row = 4, column = 0)
        
        ConfigCheck_vehicle_id_entry = Entry(ConfigCheck_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        ConfigCheck_vehicle_id_entry.grid(row = 0, column = 1)
        ConfigCheck_variant_entry = Entry(ConfigCheck_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        ConfigCheck_variant_entry.grid(row = 1, column = 1)
##        ConfigCheck_Torishi_code_entry = Entry(ConfigCheck_test_output_frame, width = 55, bd = 4,
##                                  font = ("Arial", 10, "bold"))
##        ConfigCheck_Torishi_code_entry.grid(row = 2, column = 1)
       # ConfigCheck_Torishi_code_entry = Entry(ConfigCheck_test_output_frame, width = 55, bd = 4,
        #                       font = ("Arial", 10, "bold"))
       # ConfigCheck_Torishi_code_entry.grid(row = 3, column = 1)
##        ConfigCheck_CAT_entry = Entry(ConfigCheck_test_output_frame, width = 55, bd = 4,
##                             font = ("Arial", 10, "bold"))
##        ConfigCheck_CAT_entry.grid(row = 4, column = 1)
##        
        ConfigCheck_result_entry = Entry(ConfigCheck_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        ConfigCheck_result_entry.grid(row = 3, column = 1)

        ConfigCheck_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        ConfigCheck_progressbar_style = ttk.Style()
        ConfigCheck_progressbar_style.theme_use("default")
        ConfigCheck_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = ConfigCheck_progressbar_color)
        ConfigCheck_progressbar = ttk.Progressbar(ConfigCheck_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        ConfigCheck_progressbar.grid(row = 4, column = 1)

        ConfigCheck_logging_frame = LabelFrame(frame11, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =145, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame for Config Check
        ConfigCheck_logging_frame.place(x = 505, y = 355)
        ConfigCheck_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(ConfigCheck_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 73)

        ConfigCheck_overall_progressbar = ttk.Progressbar(ConfigCheck_logging_frame,
                                              orient ="horizontal",                                                      #Creating Overall progressbar for Bus-OFF
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        ConfigCheck_overall_progressbar.place (x = 120, y = 40)
        
        testinprogress_entry = Entry(ConfigCheck_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)                      
        
        tree_frame_6 = Frame(frame11, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame_6.place(x = 15, y = 5)
        tree_frame_6.place_configure(width = 450, height = 490)

##        option_frame = Frame(background_frame,relief = RAISED, borderwidth=2,
##                           bg =background_color )                                                                     # Creating and Setting  tree frame 
##        option_frame.place(x = 15, y = 80)
##        option_frame.place_configure(width = 450, height = 180)
        

        scrollbar = Scrollbar(tree_frame_6, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        Data = ''
        tree6 = ttk.Treeview(tree_frame_6, height=25)                                                  # Tree data
        tree6.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree6.yview)
        tree6.config(yscrollcommand=scrollbar.set)
        tree6.pack(padx = 3,pady = 3)
        
          
        logo = Image.open(Image_Path+"ConfigCheck.jpg")                                                                                 # Importing the ConfigCheck logo image
        logo=logo.resize((440, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 20, y = 70)

        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)


class ManeuverTesting(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_1()

#** Definition of initializeUI (GUI creation) **#   

    def initializeUI_1(self):
        
        global tree_frame, tree, Script_Path, Org_Path
        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()

    
    

#***********************************************#       
    
        global frame12
    

        global ManeuverTesting_test_output_frame,ManeuverTesting_vehicle_id_label,ManeuverTesting_variant_label,ManeuverTesting_application_label,ManeuverTesting_testcase_label,ManeuverTesting_result_label,ManeuverTesting_Progress_label,ManeuverTesting_vehicle_id_entry,ManeuverTesting_variant_entry,ManeuverTesting_application_entry,ManeuverTesting_testcase_entry,ManeuverTesting_result_entry
        global ICC_Cancel_progressbar_color,ManeuverTesting_progressbar_style,ICC_Cancel_overall_progressbar,ICC_Cancel_progressbar
        global heading_color,Browse_button_color
                           
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        background_frame.pack(fill=BOTH, expand=True)

        heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        heading_frame = Frame(background_frame,relief=RAISED,borderwidth=3,
                              bg = heading_color)                                                                       # Creating and Setting the design of Heading Frame
        heading_frame.place(x=15, y=15)
        heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(heading_frame, image = Nissan_image, bg = heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(heading_frame, image = TCS_image, bg = heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        heading_label= Label(heading_frame, justify="center",
                             text = "ICC CANCEL Testing",                                                 # Creating the Heading Label
                             bg =heading_color, fg = "black", font="Times 25 bold")
        heading_label.pack()        
        
       
        frame12 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =browse_frame_color,height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame12.pack(fill = BOTH, expand = TRUE)
##    
       
        ManeuverTesting_test_output_frame = LabelFrame(frame12, text = "Test Case Logging",
                                       bg = browse_frame_color, width = 555,
                                       height =430, relief = RAISED, font = "Helvetica")                                 # Creating and Setting Test output label frame
        ManeuverTesting_test_output_frame.place(x = 505, y = 5)
        ManeuverTesting_test_output_frame.grid_propagate(0)

        ManeuverTesting_test_output_frame.columnconfigure(0, pad = 10)
        ManeuverTesting_test_output_frame.columnconfigure(1, pad = 10)

        ManeuverTesting_test_output_frame.rowconfigure(0, pad = 9)
        ManeuverTesting_test_output_frame.rowconfigure(1, pad = 9)
        ManeuverTesting_test_output_frame.rowconfigure(2, pad = 9)
        ManeuverTesting_test_output_frame.rowconfigure(3, pad = 9)
        ManeuverTesting_test_output_frame.rowconfigure(4, pad = 9)
        ManeuverTesting_test_output_frame.rowconfigure(5, pad = 9)
        
        logging_labels_color = '#%02x%02x%02x' % (212, 222, 222)
        ManeuverTesting_vehicle_id_label = Label(ManeuverTesting_test_output_frame,
                                 text = "Vehicle ID", width = 11, height = 2,
                                 bg = logging_labels_color, relief = RIDGE,
                                 cursor = "target", font = ("Arial", 10, "bold"))                                                  # Label creation
        ManeuverTesting_vehicle_id_label.grid(row = 0, column = 0)
        ManeuverTesting_variant_label = Label(ManeuverTesting_test_output_frame, text = "Variant", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ManeuverTesting_variant_label.grid(row = 1, column = 0)
        ManeuverTesting_application_label = Label (ManeuverTesting_test_output_frame, text = "Application",
                                   width = 11, height = 2, bg = logging_labels_color,
                                   relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ManeuverTesting_application_label.grid(row = 2, column = 0)
        ManeuverTesting_testcase_label = Label (ManeuverTesting_test_output_frame, text = "Test Case No.",
                                width =11, height = 2, bg = logging_labels_color,
                                relief = RIDGE, cursor = "target", font = ("Arial", 10, "bold"))
        ManeuverTesting_testcase_label.grid(row = 3, column = 0)
        ManeuverTesting_result_label = Label (ManeuverTesting_test_output_frame, text = "Result", width = 11,
                              height = 2, bg = logging_labels_color, relief = RIDGE,
                              cursor = "target", font = ("Arial", 10, "bold"))
        ManeuverTesting_result_label.grid(row = 4, column = 0)
        ManeuverTesting_Progress_label = Label (ManeuverTesting_test_output_frame, text = "Progress", width = 11,
                                height = 2, bg = logging_labels_color, relief = RIDGE,
                                cursor = "target", font = ("Arial", 10, "bold"))
        ManeuverTesting_Progress_label.grid(row = 5, column = 0)
        
        ManeuverTesting_vehicle_id_entry = Entry(ManeuverTesting_test_output_frame, width = 55, bd = 4,
                                 font = ("Arial", 10, "bold"))
        ManeuverTesting_vehicle_id_entry.grid(row = 0, column = 1)
        ManeuverTesting_variant_entry = Entry(ManeuverTesting_test_output_frame, width = 55, bd = 4,
                              font = ("Arial", 10, "bold"))
        ManeuverTesting_variant_entry.grid(row = 1, column = 1)
        ManeuverTesting_application_entry = Entry(ManeuverTesting_test_output_frame, width = 55, bd = 4,
                                  font = ("Arial", 10, "bold"))
        ManeuverTesting_application_entry.grid(row = 2, column = 1)
        ManeuverTesting_testcase_entry = Entry(ManeuverTesting_test_output_frame, width = 55, bd = 4,
                               font = ("Arial", 10, "bold"))
        ManeuverTesting_testcase_entry.grid(row = 3, column = 1)
        ManeuverTesting_result_entry = Entry(ManeuverTesting_test_output_frame, width = 55, bd = 4,
                             font = ("Arial", 10, "bold"))
        ManeuverTesting_result_entry.grid(row = 4, column = 1)

        ICC_Cancel_progressbar_color = '#%02x%02x%02x' % (58, 140, 44)
        ICC_Cancel_progressbar_style = ttk.Style()
        ICC_Cancel_progressbar_style.theme_use("default")
        ICC_Cancel_progressbar_style.configure("Horizontal.TProgressbar", thickness = 20,
                                    troughcolor = "white", background = ICC_Cancel_progressbar_color)
        ICC_Cancel_progressbar = ttk.Progressbar(ManeuverTesting_test_output_frame, orient ="horizontal",
                                      mode = "determinate",
                                      style = "green.Horizontal.TProgressbar",
                                      length = 250,variable = var_ind, maximum = 1)
        ICC_Cancel_progressbar.grid(row = 5, column = 1)

        
        ManeuverTesting_logging_frame = LabelFrame(frame12, text = "Overall Progress",
                                   bg = browse_frame_color, width = 555,
                                   height =130, relief = RAISED, font = "Helvetica")                                    # Creating and Setting Overall progress frame
        ManeuverTesting_logging_frame.place(x = 505, y = 365)
        ManeuverTesting_logging_frame.grid_propagate(FALSE)

                
        log_button_color = '#%02x%02x%02x' % (68, 114, 196)
        log_button = Tkinter.Button(ManeuverTesting_logging_frame, text = "Open LOG",
                                    bg = log_button_color, activebackground = "red",
                                    width = 11, bd = 3, font = "times 14 bold",
                                    cursor = "hand2", height = 1, command=Log)
        log_button.place(x = 225, y = 65)

        ICC_Cancel_overall_progressbar = ttk.Progressbar(ManeuverTesting_logging_frame,
                                              orient ="horizontal",
                                              mode = "determinate",
                                              style = "green.Horizontal.TProgressbar",
                                              length = 375,
                                              maximum = uid ,variable = var_ind)  
        ICC_Cancel_overall_progressbar.place (x = 120, y = 30)
        
        testinprogress_entry = Entry(ManeuverTesting_logging_frame, width = 40,
                                     bg = browse_frame_color, relief = FLAT,
                                     font = "times 13", justify = CENTER)
        testinprogress_entry.place(x = 105, y = 5)                     
        
        tree_frame = Frame(frame12, relief = RAISED, borderwidth=2,
                           bg = browse_frame_color)                                                                     # Creating and Setting  tree frame 
        tree_frame.place(x = 15, y = 5)
        tree_frame.place_configure(width = 450, height = 490)

        scrollbar = Scrollbar(tree_frame, bd = 3)
        scrollbar.pack(side = RIGHT, fill = Y)
        
        
        Data = ''
        tree = ttk.Treeview(tree_frame, height=25)                                                  # Tree data
        tree.column("#0",minwidth=0,width=450, stretch=NO)

        scrollbar.config(command=tree.yview)
        tree.config(yscrollcommand=scrollbar.set)
        tree.pack(padx = 3,pady = 3)
            
        ManeuverTesting_logo = Image.open(Image_Path+"ITS.JPG")                                                                          # Importing the dspace bench setup image
        ManeuverTesting_logo=ManeuverTesting_logo.resize((445, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(ManeuverTesting_logo)
        label3 = Label(self, image = logo_pi, relief = RAISED)
        label3.image = logo_pi
        label3.place(x = 20, y = 70)

        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)        
        


class ManeuverTesting2(Frame):
  
    def __init__(self, parent):
        Frame.__init__(self, parent)   
         
        self.parent = parent
        self.initializeUI_ManeuverTesting2()

                                                                    
    
    def initializeUI_ManeuverTesting2(self):
        

        var_ind = IntVar(self)
        w = 1080                                                                                                        # Width of the application window
        h = 850                                                                                                         # Height of the applicaiton window
        sw = self.parent.winfo_screenwidth()                                                                            # Width of the screen
        sh = self.parent.winfo_screenheight()                                                                           # Height of the screen
        x = (sw - w)/2                                                                                                  # X co ordinate
        y = (sh - h)/2                                                                                                  # Y co ordinate
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))                                                              # Opens the window in the center of the screen
        fp= open('HILS_Testing_Log.txt', 'w')
        fp.close()

        global frame13

       
        background_color = '#%02x%02x%02x' % (210, 210, 210)                                                            # Creating and Setting the design of background frame
        background_frame = Frame(self, relief=RAISED, borderwidth=2,
                                 bg =background_color )
        background_frame.pack(fill=BOTH, expand=True)

        heading_color = '#%02x%02x%02x' % (175, 171, 171)                                                               # Set your favourite rgb color
        heading_frame = Frame(background_frame,relief=RAISED,borderwidth=3,
                              bg = heading_color)                                                                       # Creating and Setting the design of Heading Frame
        heading_frame.place(x=15, y=15)
        heading_frame.place_configure(width = 1045, height = 50)

        Nissan_logo = Image.open(Image_Path+"Nissan_Logo.png")
        Nissan_logo = Nissan_logo.resize((55, 40), Image.ANTIALIAS)
        Nissan_image = ImageTk.PhotoImage(Nissan_logo)
        Nissan_label = Label(heading_frame, image = Nissan_image, bg = heading_color)
        Nissan_label.image = Nissan_image
        Nissan_label.place(x = 5, y = 0)
####
        TCS_logo = Image.open(Image_Path+"TCS_Logo.png")
        TCS_logo = TCS_logo.resize((50, 40), Image.ANTIALIAS)
        TCS_image = ImageTk.PhotoImage(TCS_logo)
        TCS_label = Label(heading_frame, image = TCS_image, bg = heading_color)
        TCS_label.image = TCS_image
        TCS_label.place(x = 977, y = 0)         

        heading_label= Label(heading_frame, justify="center",
                             text = "Maneuver Testing",                                                 # Creating the Heading Label
                             bg =heading_color, fg = "black", font="Times 25 bold")
        heading_label.pack()          


        frame13 = Frame(self,relief = RAISED, borderwidth = 2,
                       bg =browse_frame_color,height = 170)                                                              # Create a new frame at the bottom to accomodate test case output, tree window  
        frame13.pack(fill = BOTH, expand = TRUE) 
       

        global v11,v12,v13,v14,v15
        v11=IntVar()     # ALL button
        v12=IntVar()     #ITS button
        v13=IntVar()     #DDT button
        v14=IntVar()     #MSG_COUNTER button
        v15=IntVar()     #BSW button
       
        
        FEB_frame_seq= Frame(frame13,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for FEB button 
        FEB_frame_seq.place(x=20 ,y=20)
        FEB_frame_seq.place_configure(width = 60, height = 70)
        FEB_frame_seq_label= Label(FEB_frame_seq, justify="center",
                             text = "1.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        FEB_frame_seq_label.place(x=12,y=4)
        
        BSW_frame_seq= Frame(frame13,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for FEB button 
        BSW_frame_seq.place(x=20 ,y=100)
        BSW_frame_seq.place_configure(width = 60, height = 70)
        BSW_frame_seq_label= Label(BSW_frame_seq, justify="center",
                             text = "2.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")

        BSW_frame_seq_label.place(x=12,y=4)






        LDW_frame_seq= Frame(frame13,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for FEB button 
        LDW_frame_seq.place(x=550 ,y=20)
        LDW_frame_seq.place_configure(width = 60, height = 70)
        LDW_frame_seq_label= Label(LDW_frame_seq, justify="center",
                             text = "3.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        LDW_frame_seq_label.place(x=12,y=4)
        
        TBD_frame_seq= Frame(frame13,self,relief = RAISED, borderwidth = 4,
                           bg =frame1_color, height = 20)                                                              # Create a new frame for FEB button 
        TBD_frame_seq.place(x=550 ,y=100)
        TBD_frame_seq.place_configure(width = 60, height = 70)
        TBD_frame_seq_label= Label(TBD_frame_seq, justify="center",
                             text = "4.",                                                 # Creating the Heading Label
                              fg = "black", font="Times 25 bold")
        TBD_frame_seq_label.place(x=12,y=4)

      


        
#####################################Frames for CheckButtons#############################################################################################            

        
        FEB_frame= Frame(frame13,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for FEB button 
        FEB_frame.place(x=80 ,y=20)
        FEB_frame.place_configure(width = 390, height = 70)

        TBD_frame= Frame(frame13,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for TBD button 
        TBD_frame.place(x=610,y=100)
        TBD_frame.place_configure(width = 390, height = 70)



        BSW_frame= Frame(frame13,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for BSW button 
        BSW_frame.place(x=80,y=100)
        BSW_frame.place_configure(width = 390, height = 70)



        LDW_frame= Frame(frame13,self,relief = RAISED, borderwidth = 2,
                           bg =browse_frame_color, height = 20)                                                              # Create a new frame for LDW button 
        LDW_frame.place(x=610 ,y=20)
        LDW_frame.place_configure(width = 390, height = 70)
                         
      

        OK_frame= Frame(frame13,self,relief = RAISED, borderwidth = 2,
                           bg = browse_frame_color, height = 20)                                                              # Create a new frame for LDW button 
        OK_frame.place(x=420 ,y=220)
        OK_frame.place_configure(width = 150, height = 70)          
################################################################################################################################
        
#######################################Checknbutton Creation #####################################################################
        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)


        
        
        FEB_check_button = Tkinter.Checkbutton(FEB_frame, text = "FEB",                                #Creating FEB button
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v11,command=Assign)
        FEB_check_button.pack(side=RIGHT,anchor= NE,padx=10,pady=10)
        
        FEB_logo = Image.open(Image_Path+"FEB.JPG")
        FEB_logo = FEB_logo.resize((50, 40), Image.ANTIALIAS)
        FEB_image = ImageTk.PhotoImage(FEB_logo)
        FEB_label = Label(FEB_frame, image = FEB_image, bg = heading_color)
        FEB_label.image = FEB_image
        FEB_label.place(x = 8, y = 11)
        
        TBD_check_button = Tkinter.Checkbutton(TBD_frame, text = "T.B.D",                                    #Creating TBD check button
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v12,command=Assign)
        TBD_check_button.pack(side=RIGHT,anchor=NE,padx=10,pady=10)

        TBD_logo = Image.open(Image_Path+"TBD.jpg")
        TBD_logo = TBD_logo.resize((50, 40), Image.ANTIALIAS)
        TBD_image = ImageTk.PhotoImage(TBD_logo)
        TBD_label = Label(TBD_frame, image = TBD_image, bg = heading_color)
        TBD_label.image = TBD_image
        TBD_label.place(x = 8, y = 11)
        

        
        BSW_check_button = Tkinter.Checkbutton( BSW_frame, text = "BSW",
                                         bg=frame1_color,
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v13,command=Assign)
        BSW_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10)

        BSW_logo = Image.open(Image_Path+"BSW.JPG")
        BSW_logo = BSW_logo.resize((50, 40), Image.ANTIALIAS)
        BSW_image = ImageTk.PhotoImage(BSW_logo)
        BSW_label = Label(BSW_frame, image = BSW_image, bg = heading_color)
        BSW_label.image = BSW_image
        BSW_label.place(x = 8, y = 11)
        

        
        LDW_check_button = Tkinter.Checkbutton(LDW_frame, text = "LDW",
                                         bg=frame1_color,
                                                 
                                     width = 60,
                                     height = 3, relief = GROOVE,indicatoron=0,selectcolor=start_button_color,
                                         
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",
                                           onvalue = 1,offvalue = 0,variable=v14,command=Assign)
        LDW_check_button.pack(side=LEFT,anchor=NE,padx=10,pady=10)

        LDW_logo = Image.open(Image_Path+"LDW.jpg")
        LDW_logo = LDW_logo.resize((50, 40), Image.ANTIALIAS)
        LDW_image = ImageTk.PhotoImage(LDW_logo)
        LDW_label = Label(LDW_frame, image = LDW_image, bg = heading_color)
        LDW_label.image = LDW_image
        LDW_label.place(x = 8, y = 11)
        Browse_button_color = '#%02x%02x%02x' % (46, 117, 182)






   

        OK_button = Tkinter.Button(OK_frame, text = "OK",                                    #Creating TBD check button
                                         bg=Browse_button_color,
                                     width = 15,
                                     height = 2, relief = RAISED,
                                      borderwidth=4,   
                                     activebackground = "green", cursor = "hand2",

                                     font = "times 20 bold",command=OK_Pressed )

        #OK_button.place(x=420 ,y=360)
        OK_button.pack(padx=5,pady=5)



        ManeuverTesting_logo = Image.open(Image_Path+"BSW.JPG")                                                                          # Importing the dspace bench setup image
        ManeuverTesting_logo=ManeuverTesting_logo.resize((445, 245), Image.ANTIALIAS)
        logo_pi = ImageTk.PhotoImage(ManeuverTesting_logo)
        label3 = Label(self, image = logo_pi, relief = RAISED)
        label3.image = logo_pi
        label3.place(x = 20, y = 70)

        logo = Image.open(Image_Path+"img_to_import.jpg")                                                                          # Importing the dspace bench setup image
        logo_pi = ImageTk.PhotoImage(logo)
        label1 = Label(self, image = logo_pi, relief = RAISED)
        label1.image = logo_pi
        label1.place(x = 505, y = 70)   
      



class cWindow:
    def __init__(self):
        self._hwnd = None
        self.shell = win32com.client.Dispatch("WScript.Shell")

    def BringToTop(self):
        win32gui.BringWindowToTop(self._hwnd)
     

    def SetAsForegroundWindow(self):
        self.shell.SendKeys('%')
        win32gui.SetForegroundWindow(self._hwnd)
       

    def Maximize(self):
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)

    def setActWin(self):
        win32gui.SetActiveWindow(self._hwnd)

    def _window_enum_callback(self, hwnd, wildcard):
        '''Pass to win32gui.EnumWindows() to check all the opened windows'''
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) != None:
            self._hwnd = hwnd

    def find_window_wildcard(self, wildcard):
        self._hwnd = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)
        print wildcard
        print self._window_enum_callback
        print self._hwnd
        return self._hwnd



    def kill_task_manager(self):
        # Here I use your method to find a window because of an accent in my french OS,
        # but you should use win32gui.FindWindow(None, 'Task Manager complete name').
        wildcard = 'Gestionnaire des t.+ches de Windows'
        self.find_window_wildcard(wildcard)
        if self._hwnd:
            win32gui.PostMessage(self._hwnd, win32con.WM_CLOSE, 0, 0)  # kill it
            sleep(0.5)  # important to let time for the window to be closed   
##############################Edited################################


def main():
    global nb,app5
    root = Tkinter.Tk()
    root.title("TCS Automation Framework")

    nb = ttk.Notebook(root)
##    mainframe= Tkinter.Frame(root,name='main frame')
##    mainframe.pack(fill=Tkinter.BOTH)
##    
##    nb=ttk.Notebook(mainframe,name='nb')
##    nb.pack(fill=Tkinter.BOTH,padx=2,pady=3)
##    
##    app1=ttk.Frame(nb,name='this is 1')
##    ##    app1 = MyGUI(root)
##    Tkinter.Label(app1,text='this ias tab 1')
##    nb.add(app1,text='tab       1')
##
##    app2=ttk.Frame(nb,name='this is 2')
##    Tkinter.Label(app1,text='this ias tab 2')
##    nb.add(app2,text='tab       2')    
##    nb.select(app2)
##    page1 = ttk.Frame(nb)
##    page2 = ttk.Frame(nb)
##    page3 = ttk.Frame(nb)
##    page4 = ttk.Frame(nb)
##    page5 = ttk.Frame(nb)
##    page6 = ttk.Frame(nb)
##    page7 = ttk.Frame(nb)
##    page8 = ttk.Frame(nb)
##    
    app1 = MyGUI(root)
    app2 = ITS_Testing(root)
    app3 = Active_Testing(root)
    app4 = FAILSAFE(root)
    app5 = Message_counter(root)
    app6 = Gateway(root)
    app7 = BusOff(root)
    app8 = ConfigCheck(root)
    app10 = ManeuverTesting(root)
    app9 = ManeuverTesting2(root)
    
    #example = ThreadingExample()
    nb.add(app1, text='Automation Main Page')
    nb.add(app2, text='ITS Testing')
    nb.add(app4, text='Failsafe Testing ')
    nb.add(app5, text='Message Counter')
    nb.add(app8, text='Config Check')
    nb.add(app7, text='Bus Off')
    nb.add(app9, text='Maneuver Testing')
    nb.add(app10, text='ICC CANCEL Testing')    
    nb.add(app6, text='Gateway')
    nb.add(app3, text='Active Testing')

    
    nb.pack(expand=1, fill="both")
    root.mainloop() 


if __name__ == "__main__":
    main()
