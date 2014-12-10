VERSION 5.00
Begin VB.Form frmDigitizer_NI488 
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
      Caption         =   "Start Digitizer sample"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmDigitizer_NI488"
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
    main_Digitizer
End Sub

' This program uses the voltage in step mode and also demonstrates
' how to set up and use the digitizer.
' Add error handling and check for error after each visa command to ensure that each
' command gets send to the power supply.

' This program only works with the Agilent N675x with high-speed test extensions
' and N676x modules

Sub main_Digitizer()
    Dim SCPIcmd As String
    Dim strTemp As String * 255
    Dim IDN As String
    Dim ErrString As String
    Dim GPIBaddress As String
    Dim measPoints As Long
    Dim measOffset As Long
    Dim Voltage As Double
    Dim finalVoltage As Double
    Dim timeInterval As Double
    Dim strReadings As String
    Dim strVoltPoints() As String
    Dim VoltPoints() As Double
    Dim i As Long

    ' This variable controls the channel number to be programmed
    Dim channel As String

    'These variable are neccessary to initialize NI-488.
    Dim Address As Integer
    Dim Instrument As Integer

    Address = 5
    Instrument = ildev(0, Address, 0, T10s, 1, 0)
    If Instrument < 0 Then
        MsgBox "Unable to establish communication with the instrument"
        Exit Sub
    End If

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@1)"                                       ' channel 1

    ' Send a power reset to the instrument
    Call ibwrt(Instrument, "*RST")

    ' Query the instrument for the IDN string
    Call ibwrt(Instrument, "*IDN?")
    Call ibrd(Instrument, strTemp)
    IDN = Left$(strTemp, ibcntl - 1)

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@1)"                                       ' channel 1

    ' disable button
    Command1.Enabled = False

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@1)"                                       ' channel 1

    ' disable button
    Command1.Enabled = False

    ' This controls the number of points the measurement system measures
    measPoints = 10

    ' This controls the number of points to offset the measurement (positive for forward,
    ' negative for reverse
    measOffset = 0

    ' this sets the time between points
    timeInterval = 0.0025

    ' this variable controls the voltage
    ' it can be hard coded into the viPrintf command as well
    Voltage = 3

    ' This is the final voltage that will be triggered
    finalVoltage = 5


    ' Send a power reset to the instrument
    SCPIcmd = "*RST"
    Call ibwrt(Instrument, SCPIcmd)

    ' Query the instrument for the IDN string
    SCPIcmd = "*IDN?"
    Call ibwrt(Instrument, SCPIcmd)
    Call ibrd(Instrument, strTemp)
    IDN = Left$(strTemp, ibcntl - 1)

    ' Put the Voltage into step mode which causes it to transition from one
    ' voltage to another upon receiving a trigger
    SCPIcmd = "VOLT:MODE STEP," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' program to voltage setting
    SCPIcmd = "VOLT" & Str$(Voltage) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Go to final value
    SCPIcmd = "VOLT:TRIG" & Str$(finalVoltage) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Turn the output on
    SCPIcmd = "OUTP ON," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the bus as the transient trigger source
    SCPIcmd = "TRIG:TRAN:SOUR BUS," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the number of points for the measurement system to use as an offset
    SCPIcmd = "SENS:SWE:OFFS:POIN" & Str$(measOffset) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the number of points that the measurement system uses
    SCPIcmd = "SENS:SWE:POIN" & Str$(measPoints) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the time interval between points
    SCPIcmd = "SENS:SWE:TINT" & Str$(timeInterval) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the measurement trigger source
    SCPIcmd = "TRIG:ACQ:SOUR BUS," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Initiate the measurement trigger system
    SCPIcmd = "INIT:ACQ " & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Initiate the transient trigger system
    SCPIcmd = "INIT:TRAN " & channel
    Call ibwrt(Instrument, SCPIcmd)
    Delay 1

    ' Trigger the unit
    SCPIcmd = "*TRG"
    Call ibwrt(Instrument, SCPIcmd)
    Delay 1

    ' Read back the voltage points
    SCPIcmd = "FETC:ARR:VOLT? " & channel
    Call ibwrt(Instrument, SCPIcmd)
    Call ibrd(Instrument, strTemp)
    strReadings = Left$(strTemp, ibcntl - 1)

    ' Parse the string and put into a numeric array
    strVoltPoints = Split(strReadings, ",")
    ReDim VoltPoints(LBound(strVoltPoints) To UBound(strVoltPoints))
    For i = LBound(strVoltPoints) To UBound(strVoltPoints)
        VoltPoints(i) = Val(strVoltPoints(i))
    Next i

    ' Print the first 10 voltage points
    For i = 0 To 9
        Debug.Print i, VoltPoints(i)
    Next i

    ' Check instrument for any errors
    SCPIcmd = "Syst:err?"
    Call ibwrt(Instrument, SCPIcmd)
    Call ibrd(Instrument, strTemp)
    ErrString = Left$(strTemp, ibcntl - 1)

    ' give message if there is an error
    If Val(ErrString) Then
        MsgBox "Error in instrument!" & vbCrLf & ErrString
    End If

    ' Take the board offline
    Call ibonl(Instrument, 0)

End Sub


' Wait the specified number of seconds
Private Sub Delay(DelayTime As Single)
   Dim Finish As Single
   Finish = Timer() + DelayTime
   Do
   Loop Until Finish <= Timer()
End Sub

