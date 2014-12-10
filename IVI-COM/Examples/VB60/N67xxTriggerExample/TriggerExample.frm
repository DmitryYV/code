VERSION 5.00
Begin VB.Form TriggerExample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "N67xx Trigger Example"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "TriggerExample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
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
Attribute VB_Name = "TriggerExample"
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
' This program demonstrates how to setup a transient and measurement trigger.
'/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Option Explicit

Public driver As AgilentN67xx
Public outputPtr As IAgilentN67xxOutput3
Public measurementPtr As IAgilentN67xxMeasurement
Public transientPtr As IAgilentN67xxTransient

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdStart_Click()
Dim channel As String ' the channel to be programmed
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
        
        Set outputPtr = driver.Outputs.Item(driver.Outputs.Name(channel))
        Set measurementPtr = driver.Measurements.Item(driver.Measurements.Name(channel))
        Set transientPtr = driver.Transients.Item(driver.Transients.Name(channel))
    
        'set the voltage level
        outputPtr.VoltageLevel 51, 5
    
        ' Set the triggered voltage level and the mode
        ' Setting the voltage mode to step causes the voltage to transition from one
        ' voltage to another upon receiving a trigger
        transientPtr.StepTrigLevel(AgilentN67xxTriggerLevelVoltage) = 10

        'enable the output
        outputPtr.Enabled = True

        ' Set the transient trigger source to bus.
        transientPtr.TrigSource = AgilentN67xxTriggerSourceBus

        ' initiate the transient trigger system
        transientPtr.TrigInitiate

        'Set trigger offset in the measurement sweep
        'Set the number of data points in the measurement
        'Set the measurement sample interval
        measurementPtr.Configure AgilentN67xxMeasurementVoltage, 0, 5, 0.0025
    
        ' set the measurement trigger source
        measurementPtr.TrigSource = AgilentN67xxTriggerSourceBus

        ' initiate the measurement trigger system
        measurementPtr.TrigInitiate
        
        Delay 1
        
        ' trigger the measurement and transient systems
        driver.Transients.SendSoftwareTrigger

        ' read back the voltage points
        Dim fetchArray() As Double
        fetchArray = measurementPtr.fetchArray(AgilentN67xxFetchVoltage)

        ' Print the voltage points
        Dim str As String
        Dim i As Integer
        For i = 0 To 4
            str = str & fetchArray(i)
            If i <> 4 Then
               str = str & ", "
            End If
        Next i
    
        StatusTextBox.Text = StatusTextBox.Text & "The measurement results are : " & str
        StatusTextBox.Refresh
        
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

' Wait the specified number of seconds
Private Sub Delay(DelayTime As Single)
   Dim Finish As Single
   Finish = Timer() + DelayTime
   Do
   Loop Until Finish <= Timer()
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

