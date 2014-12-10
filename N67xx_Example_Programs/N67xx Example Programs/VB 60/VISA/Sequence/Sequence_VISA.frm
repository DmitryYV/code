VERSION 5.00
Begin VB.Form frmSequence_VISA 
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
      Caption         =   "Start Sequence sample"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmSequence_VISA"
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
    main_Sequence
End Sub

' This program sets up two output channels to turn on in sequence
' The first ouput channel turns on when the OUTP:ON command is received
' The second turns on 50 ms later
Sub main_Sequence()
    Dim SCPIcmd As String
    Dim actual As Long
    Dim strTemp As String * 255
    Dim chA As String
    Dim chB As String
    Dim chAB As String
    Dim voltA As Double
    Dim voltB As Double
    Dim currA As Double
    Dim currB As Double
    Dim onDelayB As Double
    Dim ErrString As String
    Dim GPIBaddress As String
    
    
    ' The two channels to be programmed are 1 and 2
    chA = "(@1)"
    chB = "(@2)"
    chAB = "(@1,2)"
    
    ' Channel one at 3 V
    ' Channel two at 5 V
    voltA = 3    ' in volts
    voltB = 5   ' in volts
    
    ' Current on channel one is 0.3 A
    ' Current on channel two is 0.5 A
    currA = 0.3   ' in amps
    currB = 0.5   ' in amps
     
    ' There will be a 50 ms delay between the two channels
    ' turning on
    onDelayB = 0.05 ' in seconds
    
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
    
    ' use the following line instead for USB communication
    ' USBaddress = "USB0::2391::1799::US00000002"

    status = viOpenDefaultRM(viDefaultRM)
    status = viOpen(viDefaultRM, GPIBaddress, 0, 2500, Instrument)
    If status < 0 Then
        MsgBox "Unable to open Port"
        Exit Sub
    End If

    ' Set voltage and current for both channels
    SCPIcmd = "VOLT" & Str$(voltA) & "," & chA
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    SCPIcmd = "VOLT" & Str$(voltB) & "," & chB
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    
    ' Set current level
    SCPIcmd = "CURR " & Str$(currA) & "," & chA
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    SCPIcmd = "CURR " & Str$(currB) & "," & chB
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)
    
    ' This command sets the output on delay
    ' An output off delay can also be set
    SCPIcmd = "Outp:Del:Rise  " & Str$(onDelayB) & "," & chB
    Call viWrite(Instrument, SCPIcmd & Chr$(10), Len(SCPIcmd) + 1, actual)

    ' Turn the output on
    SCPIcmd = "OUTP ON," & chAB
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

