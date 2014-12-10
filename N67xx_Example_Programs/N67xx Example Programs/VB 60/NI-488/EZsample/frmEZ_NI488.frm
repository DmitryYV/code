VERSION 5.00
Begin VB.Form frmEZ_NI488 
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
Attribute VB_Name = "frmEZ_NI488"
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
    Dim strTemp As String * 255
    Dim IDN As String
    Dim ErrString As String
    
    ' This variable controls the channel number to be programmed
    Dim channel As String
    
    ' This variable controls the voltage
    Dim VoltSetting As Double
    
    ' This variable controls the current
    Dim CurrSetting As Double
    
    ' These variables control the over voltage protection settings
    Dim overVoltOn As Long
    Dim overVoltSetting As Double
    
    ' These variables control the over current protection
    Dim overCurrentOn As Long
    Dim ocercurrentSetting As Double
    
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
    channel = "(@1)"        ' channel 1
    
    VoltSetting = 3
    CurrSetting = 0.5       ' amps
    overVoltSetting = 5
    

    ' Send a power reset to the instrument
    Call ibwrt(Instrument, "*RST")

    ' Query the instrument for the IDN string
    Call ibwrt(Instrument, "*IDN?")
    Call ibrd(Instrument, strTemp)
    IDN = Left$(strTemp, ibcntl - 1)

    ' Set voltage
    SCPIcmd = "VOLT" & Str$(VoltSetting) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the over voltage level
    SCPIcmd = "VOLT:PROT:LEV " & Str$(overVoltSetting) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set current level
    SCPIcmd = "CURR " & Str$(CurrSetting) & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Turn the output on
    SCPIcmd = "OUTP ON," & channel
    Call ibwrt(Instrument, SCPIcmd)

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

