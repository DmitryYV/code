VERSION 5.00
Begin VB.Form frmEZ_VisaCom 
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
Attribute VB_Name = "frmEZ_VisaCom"
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

' This is a simple program that sets a voltage, current, overvoltage, the
' status of over current protection, turns the output on and takes a
' voltage measurement.
Sub main_EZ()
    Dim IDN As String
    Dim GPIBaddress As String
    Dim ErrString As String
    
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
    
    ' This variable will contain the voltage measurement
    Dim reading As Double

    ' These variables control the over current protection
    Dim overCurrentOn As Long
    Dim ocercurrentSetting As Double

    'These variable are neccessary to initialize the VISA COM.
    Dim ioMgr As AgilentRMLib.SRMCls
    Dim Instrument As VisaComLib.FormattedIO488

    ' The following command line provides the program with the VISA name of the
    ' interface that it will be communication with.
    ' It is currently set to use GPIB to communicate
    GPIBaddress = "GPIB0::5::INSTR"

    ' Use the following line for LAN communication
    ' TCPIPaddress="TCPIP0::141.25.36.214"

    ' use the following line instead for USB communication
    ' USBaddress = "USB0::2391::1799::US00000002"

    ' Initialize the VISA COM communication
    Set ioMgr = New AgilentRMLib.SRMCls
    Set Instrument = New VisaComLib.FormattedIO488
    Set Instrument.IO = ioMgr.Open(GPIBaddress)

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@1)"                                       ' channel 1

    VoltSetting = 3
    CurrSetting = 0.5                                      ' amps
    overVoltSetting = 5.1

    With Instrument
        ' Send a power reset to the instrument
        .WriteString "*RST"

        ' Query the instrument for the IDN string
        .WriteString "*IDN?"
        IDN = .ReadString


        ' Set voltage
        .WriteString "VOLT" & Str$(VoltSetting) & "," & channel

        ' Set the over voltage level
        .WriteString "VOLT:PROT:LEV " & Str$(overVoltSetting) & "," & channel

        ' Turn on over current protection
        .WriteString "CURR:PROT:STAT ON," & channel

        ' Set current level
        .WriteString "CURR " & Str$(CurrSetting) & "," & channel

        ' Turn the output on
        .WriteString "OUTP ON," & channel
        
        Delay 1
        ' Measure the voltage
        .WriteString "Meas:Volt? " & channel
        MeasureVoltString = .ReadNumber
        MsgBox "Measured Voltage is " & MeasureVoltString & " at channel" & channel

        ' Check instrument for any errors
        .WriteString "Syst:err?"
        ErrString = .ReadString

        ' give message if there is an error
        If Val(ErrString) Then
            MsgBox "Error in instrument!" & vbCrLf & ErrString
        End If
    End With

End Sub

' Wait the specified number of seconds
Private Sub Delay(DelayTime As Single)
   Dim Finish As Single
   Finish = Timer() + DelayTime
   Do
   Loop Until Finish <= Timer()
End Sub

