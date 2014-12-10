VERSION 5.00
Begin VB.Form frmList_NI488 
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
      Caption         =   "Start List sample"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmList_NI488"
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
    main_List
End Sub

' This program executes a 10 point current and voltage list.  It also specifies
' 10 different dwell times. For ease in reading error checking is not included
' Add error handling and check for error after each visa command to ensure that each
' command gets send to the power supply.
'
' This program only works with the Agilent N675x with high-speed test extensions
' and N676x modules specified for the voltage and current ranges embedded in the
' code below.
Sub main_List()
    Dim SCPIcmd As String
    Dim strTemp As String * 255
    Dim IDN As String
    Dim ErrString As String

    ' This variable controls the channel number to be programmed
    Dim channel As String

    ' These next three strings are the points in the list.
    ' All three strings are the same length.
    ' The first one controls voltage, the second current, and the third dwell time
    Const voltPoints = "1,2,3,4,5,6,7,8,9,10"
    Const currPoints = "0.5,1,1.5,2,2.5,3,3.5,4,4.5,5"
    Const dwellPoints = "1,2,0.5,1,0.25,1.5,0.1,1,0.75,1.2"

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
    channel = "(@3)"              ' channel 1


    ' Send a power reset to the instrument
    Call ibwrt(Instrument, "*RST")

    ' Query the instrument for the IDN string
    Call ibwrt(Instrument, "*IDN?")
    Call ibrd(Instrument, strTemp)
    IDN = Left$(strTemp, ibcntl - 1)

    ' Set the voltage mode to list
    SCPIcmd = "VOLT:MODE LIST," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the current mode to list
    SCPIcmd = "CURR:MODE LIST," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Send the voltage list points
    SCPIcmd = "LIST:VOLT " & voltPoints & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Send the Current list points
    SCPIcmd = "LIST:CURR " & currPoints & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Send the dwell points
    SCPIcmd = "LIST:DWEL " & dwellPoints & "," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Turn the output on
    SCPIcmd = "OUTP ON," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Set the trigger source to bus
    SCPIcmd = "TRIG:TRAN:SOUR BUS," & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Initiate the transient system
    SCPIcmd = "INIT:TRAN " & channel
    Call ibwrt(Instrument, SCPIcmd)

    ' Trigger the unit
    SCPIcmd = "*TRG"
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

