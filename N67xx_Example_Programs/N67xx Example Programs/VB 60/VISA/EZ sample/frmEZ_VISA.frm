VERSION 5.00
Begin VB.Form frmEZ_VISA 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start EZ sample"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmEZ_VISA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
''  © Agilent Technologies, Inc. 2007
''
'' You have a royalty-free right to use, modify, reproduce and distribute
'' the Sample Application Files (and/or any modified version) in any way
'' you find useful, provided that you agree that Agilent Technologies has no
'' warranty,  obligations or liability for any Sample Application Files.
''
'' Agilent Technologies provides programming examples for illustration only,
'' This sample program assumes that you are familiar with the programming
'' language being demonstrated and the tools used to create and debug
'' procedures. Agilent Technologies support engineers can help explain the
'' functionality of Agilent Technologies software components and associated
'' commands, but they will not modify these samples to provide added
'' functionality or construct procedures to meet your specific needs.
'' """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Private Sub Command1_Click()
    main_EZ
End Sub

' This is a simple program that sets a voltage, current, overvoltage, and the
' status of over current protection. ' 10 different dwell times.
' Add error handling and check for error after each visa command to ensure that each
' command gets send to the power supply.
Sub main_EZ()
    Dim SCPIcmd As String
    Dim actual As Long
    Dim strTemp As String * 255
    Dim IDN As String
    Dim ErrString As String
    Dim GPIBaddress As String
    
    ' This variable controls the channel number to be programmed
    Dim channel As String
    
    ' This variable controls the voltage
    Dim VoltSetting As Double
    
    ' This variable reading the voltage
    Dim MeasureVoltString As String
    
    ' This variable controls the current
    Dim CurrSetting As Double
    
    ' These variables control the over voltage protection settings
    Dim overVoltOn As Long
    Dim overVoltSetting As Double
    
    ' These variables control the over current protection
    Dim overCurrentOn As Long
    Dim ocercurrentSetting As Double
    
    'These variable are neccessary to initialize the VISA system.
    Dim status As Long
    Dim viDefaultRM As Long
    Dim Instrument As Long

    ' The following command line provides the program with the VISA name of the
    ' interface that it will be communication with.
    ' It is currently set to use GPIB to communicate
    GPIBaddress = "GPIB0::5::INSTR"
    
    ' Use the following line for LAN communication
    ' TCPIPaddress="TCPIP0::141.25.36.214"
    ' TCPIP0 is the VISA name of the interface
    ' 141.25.36.214 is the IP address of the instrument
    
    ' use the following line instead for USB communication
    ' USBaddress = "USB0::2391::1799::US00000002"
    ' USB0 is the VISA name of the interface
    ' 2391 is a unique identifier for Agilent
    ' 1799 is the instrument identifier
    ' US00000002 is the instrument serial number

    status = viOpenDefaultRM(viDefaultRM)
    status = viOpen(viDefaultRM, GPIBaddress, 0, 2500, Instrument)
    If status < 0 Then
        MsgBox "Unable to open Port"
        Exit Sub
    End If

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@1)"        ' channel 1
    
    VoltSetting = 3
    CurrSetting = 0.5       ' amps
    overVoltSetting = 5.1
    

    ' Send a power reset to the instrument
    SCPIcmd = "*RST"
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Query the instrument for the IDN string
    SCPIcmd = "*IDN?"
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    Call viRead(Instrument, strTemp, 255, actual)
    IDN = Left$(strTemp, actual - 1)

    ' Set voltage
    SCPIcmd = "VOLT" & Str$(VoltSetting) & "," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Set the over voltage level
    SCPIcmd = "VOLT:PROT:LEV " & Str$(overVoltSetting) & "," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Set current level
    SCPIcmd = "CURR " & Str$(CurrSetting) & "," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Turn the output on
    SCPIcmd = "OUTP ON," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    Delay 1
    ' Measure the voltage
    SCPIcmd = "MEAS:VOLT? " & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    Call viRead(Instrument, strTemp, 255, actual)
    MeasureVoltString = Left$(strTemp, actual - 1)
    MsgBox "Measured Voltage is " & MeasureVoltString & " at channel" & channel
    
    ' Check instrument for any errors
    SCPIcmd = "Syst:err?"
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    Call viRead(Instrument, strTemp, 255, actual)
    ErrString = Left$(strTemp, actual - 1)
    
    ' give message if there is an error
    If Val(ErrString) Then
        MsgBox "Error in instrument!" & vbCrLf & ErrString
    End If
    
    ' Free up all of the resources that were assigned to the VISA system.
    ' frees up any memory that was being taken up
    status = viClose(Instrument)
    status = viClose(viDefaultRM)
    
End Sub

' Wait the specified number of seconds
Private Sub Delay(DelayTime As Single)
   Dim Finish As Single
   Finish = Timer() + DelayTime
   Do
   Loop Until Finish <= Timer()
End Sub
