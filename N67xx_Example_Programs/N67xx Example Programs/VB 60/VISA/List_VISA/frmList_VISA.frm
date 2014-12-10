VERSION 5.00
Begin VB.Form frmList_VISA 
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
      Caption         =   "Start List"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmList_VISA"
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
    Dim actual As Long
    Dim strTemp As String * 255
    Dim IDN As String
    Dim ErrString As String
    Dim GPIBaddress As String
    ' This variable controls the channel number to be programmed
    Dim channel As String
    'These variable are neccessary to initialize the VISA system.
    Dim status As Long
    Dim viDefaultRM As Long
    Dim Instrument As Long


    ' These next three strings are the points in the list.
    ' All three strings are the same length.
    ' The first one controls voltage, the second current, and the third dwell time
    Const voltPoints = "1,2,3,4,5,6,7,8,9,10"
    Const currPoints = "0.5,1,1.5,2,2.5,3,3.5,4,4.5,5"
    Const dwellPoints = "1,2,0.5,1,0.25,1.5,0.1,1,0.75,1.2"


    ' The following command line provides the program with the VISA name of the
    ' interface that it will be communication with.
    ' It is currently set to use GPIB to communicate
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
    
    GPIBaddress = "GPIB0::5::INSTR"

    status = viOpenDefaultRM(viDefaultRM)
    status = viOpen(viDefaultRM, GPIBaddress, 0, 2500, Instrument)
    If status < 0 Then
        MsgBox "Unable to open Port"
        Exit Sub
    End If

    ' This variable can be changed to program any channel in the mainframe
    channel = "(@3)"

    ' Send a power reset to the instrument
    SCPIcmd = "*RST"
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Query the instrument for the IDN string
    SCPIcmd = "*IDN?"
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    Call viRead(Instrument, strTemp, 255, actual)
    IDN = Left$(strTemp, actual - 1)

    ' Set the voltage mode to list
    SCPIcmd = "VOLT:MODE LIST," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Set the current mode to list
    SCPIcmd = "CURR:MODE LIST," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Send the voltage list points
    SCPIcmd = "LIST:VOLT " & voltPoints & "," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Send the Current list points
    SCPIcmd = "LIST:CURR " & currPoints & "," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Send the dwell points
    SCPIcmd = "LIST:DWEL " & dwellPoints & "," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Turn the output on
    SCPIcmd = "OUTP ON," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Set the trigger source to bus
    SCPIcmd = "TRIG:TRAN:SOUR BUS," & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Initiate the transient system
    SCPIcmd = "INIT:TRAN " & channel
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    
    ' Trigger the unit
    SCPIcmd = "*TRG"
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    
    
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

