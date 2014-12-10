VERSION 5.00
Begin VB.Form ListExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Programming Example"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "N67xxListExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "N67xxListExample.frx":014A
   ScaleHeight     =   3570
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Example"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Example"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox StatusTextBox 
      Height          =   2895
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   7455
   End
End
Attribute VB_Name = "ListExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/****************************************************************************
    '    Copyright © 2003-05 Agilent Technologies Inc. All rights
    '    reserved.
    '
    ' You have a royalty-free right to use, modify, reproduce and distribute
    ' the Sample Application Files (and/or any modified version) in any way
    ' you find useful, provided that you agree that Agilent has no
    ' warranty,  obligations or liability for any Sample Application Files.
    '
    ' Agilent Technologies provides programming examples for illustration only,
    ' This sample program assumes that you are familiar with the programming
    ' language being demonstrated and the tools used to create and debug
    ' procedures. Agilent support engineers can help explain the
    ' functionality of Agilent software components and associated
    ' commands, but they will not modify these samples to provide added
    ' functionality or construct procedures to meet your specific needs.
 '****************************************************************************/
'/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'This program executes a 3 point current and voltage list.  It also specifies
'3 different dwell times.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
Option Explicit

Public driver As AgilentN67xx
Public outputPtr As IAgilentN67xxOutput3
Public transientPtr As IAgilentN67xxTransient

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdStart_Click()
   Dim channel As String  ' the channel to be programmed
    channel = 1
    
    Set driver = New AgilentN67xx
    
    On Error GoTo errorHandler
   
    
    ' Get driver Identity properties.  Driver initialization not required.
    StatusTextBox.Text = ""
    StatusTextBox.Text = StatusTextBox.Text & "Identifier: " & driver.Identity.Identifier
    StatusTextBox.Text = StatusTextBox.Text & "   Revision: " & driver.Identity.Revision & vbCrLf
    StatusTextBox.Text = StatusTextBox.Text & "Description: " & driver.Identity.Description & vbCrLf
    StatusTextBox.Text = StatusTextBox.Text & "Initializing... "
    StatusTextBox.Refresh
    
        ' initialize the driver
    
        driver.Initialize "GPIB0::5::INSTR", _
                            False, _
                            True, _
                            "Cache=true,InterchangeCheck=false,QueryInstrStatus=true,Simulate=false"
                            
        Dim result As Boolean
        result = driver.Initialized
        
        StatusTextBox.Text = StatusTextBox.Text & "Done." & vbCrLf
        StatusTextBox.Refresh
        
        ' get references to the needed interfaces
        Set outputPtr = driver.Outputs.Item(driver.Outputs.Name(channel))
        Set transientPtr = driver.Transients.Item(driver.Transients.Name(channel))
      
        ' create arrays for the ListPoints method
        Dim voltList(2) As Double
        Dim currList(2) As Double
        Dim dwellTime(2) As Double
        
        voltList(0) = 5
        voltList(1) = 7
        voltList(2) = 10
        
        currList(0) = 0.25
        currList(1) = 0.5
        currList(2) = 1#
        
        dwellTime(0) = 1
        dwellTime(1) = 2
        dwellTime(2) = 0.5

        ' call ListPoints to set the list values and set the voltage and current modes to LIST
        transientPtr.ListPoints voltList, currList, dwellTime
        
        ' enable the output
        outputPtr.Enabled = True

        ' Set the transient trigger source to bus.
        transientPtr.TrigSource = AgilentN67xxTriggerSourceBus
        
        ' initiate the transient trigger system
        transientPtr.TrigInitiate
        
        Delay (1)
                
        ' trigger the transient system
        driver.Transients.SendSoftwareTrigger
        
        ReadInstrumentError driver
                      
        driver.Close
        
        StatusTextBox.Text = StatusTextBox.Text & "Driver closed." & vbCrLf
        StatusTextBox.Refresh
    Exit Sub
errorHandler:
    MsgBox Err.Description
    driver.Close
    Exit Sub
End Sub

Private Sub ReadInstrumentError(agDrvr As IIviDriver)
    ' Read instrument error queue until its empty.
    Dim errCode As Long
    errCode = 999
    Dim errMsg As String
    
    StatusTextBox.Text = StatusTextBox.Text & vbCrLf
    While errCode <> 0
        agDrvr.Utility.ErrorQuery errCode, errMsg
        StatusTextBox.Text = StatusTextBox.Text & "ErrorQuery: " & errCode & ", " & errMsg & vbCrLf
    Wend
End Sub

' Wait the specified number of seconds
Private Sub Delay(DelayTime As Single)
   Dim Finish As Single
   Finish = Timer() + DelayTime
   Do
   Loop Until Finish <= Timer()
End Sub

