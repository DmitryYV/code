'/****************************************************************************
'    Copyright © 2003-07 Agilent Technologies Inc. All rights
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
'This program executes a 3 point current and voltage list.  It also specifies
'3 different dwell times.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

Module Module1

    ' Agilent IVI-COM Driver Example Program
    ' 
    ' Creates a driver object and reads the Identifier and
    ' InstrumentFirmwareRevision properties.
    ' You must add a COM reference to the driver's type library.
    Private driver As Agilent.AgilentN67xx.Interop.AgilentN67xx
    Public outputPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxOutput3
    Public transientPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxTransient

    Sub Main()
        ' Create an instance of the driver
        Dim driver As New Agilent.AgilentN67xx.Interop.AgilentN67xx
        Dim channel As String  ' the channel to be programmed
        channel = 1

        Try
            ' Get Identifier property.  Driver initialization not required.
            Console.Write("Identifier: " & driver.Identity.Identifier & vbCrLf)
            Console.Write("Revision: " & driver.Identity.Revision & vbCrLf)
            Console.Write("Description: " & driver.Identity.Description & vbCrLf)
            Console.Write("Initializing... ")

            ' Initialize driver 
            driver.Initialize("GPIB0::5::INSTR", _
                            False, _
                            True, _
                            "Cache=true,InterchangeCheck=false,QueryInstrStatus=true,Simulate=false") 'Optional ivi options

            Dim result As Boolean
            result = driver.Initialized

            Console.WriteLine("Done." & vbCrLf)

            ' Get InstrumentModel and InstrumentFirmwareRevision property.
            Console.Write("InstrumentModel:  " + driver.Identity.InstrumentModel & vbCrLf)
            Console.Write("FirmwareRevision: " + driver.Identity.InstrumentFirmwareRevision & vbCrLf)

            ' get references to the needed interfaces
            outputPtr = driver.Outputs.Item(driver.Outputs.Name(channel))
            transientPtr = driver.Transients.Item(driver.Transients.Name(channel))

            ' create arrays for the ListPoints method
            Dim voltList(2) As Double
            Dim currList(2) As Double
            Dim dwellTime(2) As Double

            voltList(0) = 5
            voltList(1) = 7
            voltList(2) = 10

            currList(0) = 0.25
            currList(1) = 0.5
            currList(2) = 1.0#

            dwellTime(0) = 1
            dwellTime(1) = 2
            dwellTime(2) = 0.5

            ' call ListPoints to set the list values and set the voltage and current modes to LIST
            transientPtr.ListPoints(voltList, currList, dwellTime)

            ' enable the output
            outputPtr.Enabled = True

            ' Set the transient trigger source to bus.
            transientPtr.TrigSource = Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerSourceEnum.AgilentN67xxTriggerSourceBus

            ' initiate the transient trigger system
            transientPtr.TrigInitiate()

            ' trigger the transient system
            driver.Transients.SendSoftwareTrigger()

            driver.Systems.WaitForOperationComplete(5000)

            ReadInstrumentError(driver)

            ' Close driver if initialized.
            If (driver.Initialized) Then driver.Close()

            Console.WriteLine()
            Console.Write("Press Enter to Exit ")
            Console.ReadLine()

        Catch err As System.Exception
            Console.WriteLine()
            Console.WriteLine("Exception Error:")
            Console.WriteLine("  " + err.Message())

            driver.Close()

            Console.WriteLine()
            Console.Write("Press Enter to Exit ")
            Console.ReadLine()
        End Try


    End Sub
    Private Sub ReadInstrumentError(ByVal agDrvr As Agilent.AgilentN67xx.Interop.AgilentN67xx)
        ' Read instrument error queue until its empty.
        Dim errCode As Long
        errCode = 999
        Dim errMsg As String

        While errCode <> 0
            agDrvr.Utility.ErrorQuery(errCode, errMsg)
            Console.WriteLine(vbCrLf & "ErrorQuery: " & errCode & ", " & errMsg)
        End While
    End Sub

End Module

