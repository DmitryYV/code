VERSION 5.00
Begin VB.Form frmDigitizer_VisaCom 
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
Attribute VB_Name = "frmDigitizer_VisaCom"
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

' This program only works with the Agilent N675x with high-speed test extensions
' and N676x modules

Sub main_Digitizer()
    Dim IDN As String
    Dim GPIBaddress As String
    Dim ErrString As String
    Dim channel As String
    Dim measPoints As Long
    Dim measOffset As Long
    Dim Voltage As Double
    Dim finalVoltage As Double
    Dim timeInterval As Double
    Dim VoltPoints() As Variant
    Dim i As Long

    'These variable are neccessary to initialize the VISA COM.
    Dim ioMgr As AgilentRMLib.SRMCls
    Dim Instrument As VisaComLib.FormattedIO488

    ' disable button
    Command1.Enabled = False

    ' The following command line provides the program with the VISA name of the
    ' interface that it will be communication with.
    ' It is currently set to use GPIB to communicate
    GPIBaddress = "GPIB0::5::INSTR"

    ' Use the following line for LAN communication
    ' GPIBaddress="TCPIP0::141.25.36.214"

    ' use the following line instead for USB communication
    ' GPIBaddress = "USB0::"


    ' This controls the number of points the measurement system measures
    measPoints = 100

    ' This controls the number of points to offset the measurement (positive for forward,
    ' negative for reverse
    measOffset = 0

    ' this sets the time between points
    timeInterval = 0.0025

    ' this variable controls the voltage
    ' it can be hard coded into the viPrintf command as well
    Voltage = 5

    ' This is the final voltage that will be triggered
    finalVoltage = 10

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@1)"                                       ' channel 1


    ' Initialize the VISA COM communication
    Set ioMgr = New AgilentRMLib.SRMCls
    Set Instrument = New VisaComLib.FormattedIO488
    Set Instrument.IO = ioMgr.Open(GPIBaddress)


    With Instrument
        ' Send a power reset to the instrument
        .WriteString "*RST"

        ' Query the instrument for the IDN string
        .WriteString "*IDN?"
        IDN = .ReadString

        ' Put the Voltage into step mode which causes it to transition from one
        ' voltage to another upon receiving a trigger
        .WriteString "VOLT:MODE STEP," & channel

        ' program to voltage setting
        .WriteString "VOLT" & Str$(Voltage) & "," & channel

        ' Go to final value
        .WriteString "VOLT:TRIG" & Str$(finalVoltage) & "," & channel

        ' Turn the output on
        .WriteString "OUTP ON," & channel

        ' Set the bus as the transient trigger source
        .WriteString "TRIG:TRAN:SOUR BUS," & channel

        ' Set the number of points for the measurement system to use as an offset
        .WriteString "SENS:SWE:OFFS:POIN" & Str$(measOffset) & "," & channel

        ' Set the number of points that the measurement system uses
        .WriteString "SENS:SWE:POIN" & Str$(measPoints) & "," & channel

        ' Set the time interval between points
        .WriteString "SENS:SWE:TINT" & Str$(timeInterval) & "," & channel

        ' Set the measurement trigger source
        .WriteString "TRIG:ACQ:SOUR BUS," & channel

        ' Initiate the measurement trigger system
        .WriteString "INIT:ACQ " & channel

        ' Initiate the transient trigger system
        .WriteString "INIT:TRAN " & channel
        Delay 1

        ' Trigger the unit
        .WriteString "*TRG"
        Delay 1
        
        ' Read back the voltage points
        .WriteString "FETC:ARR:VOLT? " & channel
        VoltPoints = .ReadList
        
        ' Print the first 10 voltage points
        For i = 0 To 9
            Debug.Print i, VoltPoints(i)
        Next i

        ' Check instrument for any errors
        .WriteString "Syst:err?"
        ErrString = .ReadString

        ' give message if there is an error
        If Val(ErrString) Then
            MsgBox "Error in instrument!" & vbCrLf & ErrString
        End If
    End With

    Command1.Enabled = True

End Sub

' Wait the specified number of seconds
Private Sub Delay(DelayTime As Single)
   Dim Finish As Single
   Finish = Timer() + DelayTime
   Do
   Loop Until Finish <= Timer()
End Sub

