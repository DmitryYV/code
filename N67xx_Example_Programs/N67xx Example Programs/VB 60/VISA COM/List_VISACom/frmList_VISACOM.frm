VERSION 5.00
Begin VB.Form frmList_VisaCom 
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
Attribute VB_Name = "frmList_VisaCom"
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
'
' This program only works with the Agilent N675x with high-speed test extensions
' and N676x modules specified for the voltage and current ranges embedded in the
' code below.

Sub main_List()
    Dim IDN As String
    Dim GPIBaddress As String
    Dim ErrString As String
    Dim channel As String
    
    ' These next three strings are the points in the list.
    ' All three strings are the same length.
    ' The first one controls voltage, the second current, and the third dwell time
    Const voltPoints = "1,2,3,4,5,6,7,8,9,10"
    Const currPoints = "0.5,1,1.5,2,2.5,3,3.5,4,4.5,5"
    Const dwellPoints = "1,2,0.5,1,0.25,1.5,0.1,1,0.75,1.2"

    'These variable are neccessary to initialize the VISA COM.
    Dim ioMgr As AgilentRMLib.SRMCls
    Dim Instrument As VisaComLib.FormattedIO488
    Dim ioAddress As String

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
    channel = "(@3)"                                       ' channel 1


    With Instrument
        ' Send a power reset to the instrument
        .WriteString "*RST"

        ' Query the instrument for the IDN string
        .WriteString "*IDN?"
        IDN = .ReadString


        ' Set the voltage mode to list
        .WriteString "VOLT:MODE LIST," & channel

        ' Set the current mode to list
        .WriteString "CURR:MODE LIST," & channel

        ' Send the voltage list points
        .WriteString "LIST:VOLT " & voltPoints & "," & channel

        ' Send the Current list points
        .WriteString "LIST:CURR " & currPoints & "," & channel

        ' Send the dwell points
        .WriteString "LIST:DWEL " & dwellPoints & "," & channel

        ' Turn the output on
        .WriteString "OUTP ON," & channel

        ' Set the trigger source to bus
        .WriteString "TRIG:TRAN:SOUR BUS," & channel

        ' Initiate the transient system
        .WriteString "INIT:TRAN " & channel

        ' Trigger the unit
        .WriteString "*TRG"

        ' Check instrument for any errors
        .WriteString "Syst:err?"
        ErrString = .ReadString

        ' give message if there is an error
        If Val(ErrString) Then
            MsgBox "Error in instrument!" & vbCrLf & ErrString
        End If
    End With

End Sub

