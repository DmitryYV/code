VERSION 5.00
Begin VB.Form N67xxOutputExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "N67xx Output Example"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "OutputExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StatusTextBox 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   7455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Example"
      Height          =   375
      Left            =   5850
      TabIndex        =   1
      Top             =   3090
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Example"
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   3090
      Width           =   1695
   End
End
Attribute VB_Name = "N67xxOutputExample"
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
' This program demonstrates how to communicate with a N67xx Modular Power System
' using the IVI-COM driver.
'/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit

Public driver As AgilentN67xx
Public outputPtr As IAgilentN67xxOutput3
Public protectionPtr As IAgilentN67xxProtection2
Public measurementPtr As IAgilentN67xxMeasurement

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdStart_Click()
        Dim channel As String  ' the channel to be programmed
        channel = 1
          
        Set driver = New AgilentN67xx
          
        On Error GoTo ErrorHandler
           
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
        
        Set outputPtr = driver.Outputs.Item(driver.Outputs.Name(channel))
        Set protectionPtr = driver.Protections.Item(driver.Protections.Name(channel))
        Set measurementPtr = driver.Measurements.Item(driver.Measurements.Name(channel))
    
        ' set voltage level
        outputPtr.VoltageLevel 3#, 3#
        
        ' enable OV protection and set OV level
        protectionPtr.ConfigureOVP 10
        
        'enable the over current protection
        protectionPtr.CurrentLimitBehavior = AgilentN67xxCurrentLimitTrip
        
        ' set current level
        outputPtr.CurrentLimit 1#, 1#
        
        ' enable the output
        outputPtr.Enabled = True
                        
        Delay 1
        
        ' Measure the voltage
        Dim measVoltage As Double
        measVoltage = measurementPtr.Measure(AgilentN67xxMeasurementVoltage)
        
        ' display the measured voltage
        StatusTextBox.Text = StatusTextBox.Text & "Measured Voltage is " & measVoltage & " at channel " & channel
        StatusTextBox.Refresh

         ReadInstrumentError driver
    
        driver.Close
        
         StatusTextBox.Text = StatusTextBox.Text & "Driver closed." & vbCrLf
        StatusTextBox.Refresh
Exit Sub
ErrorHandler:
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
