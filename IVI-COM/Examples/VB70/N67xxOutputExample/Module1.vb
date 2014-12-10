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
' This program demonstrates how to communicate with a N67xx Modular Power System
' using the IVI-COM driver.
'/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Module Module1
    ' Agilent IVI-COM Driver Example Program
    ' 
    ' Creates a driver object and reads the Identifier and
    ' InstrumentFirmwareRevision properties.
    ' You must add a COM reference to the driver's type library.
    Private driver As Agilent.AgilentN67xx.Interop.AgilentN67xx
    Public outputPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxOutput3
    Public protectionPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxProtection2
    Public measurementPtr As Agilent.AgilentN67xx.Interop.AgilentN67xxMeasurement

    Sub Main()
        Dim channel As String  ' the channel to be programmed
        channel = 1

        Dim driver As New Agilent.AgilentN67xx.Interop.AgilentN67xx

        Try

            ' Get driver Identity properties.  Driver initialization not required.

            Console.Write("Identifier: " & driver.Identity.Identifier & vbCrLf)
            Console.Write("Revision: " & driver.Identity.Revision & vbCrLf)
            Console.Write("Description: " & driver.Identity.Description & vbCrLf)
            Console.Write("Initializing... ")

            ' initialize the driver
            driver.Initialize("GPIB0::5::INSTR", _
                                    False, _
                                    True, _
                                    "Cache=true,InterchangeCheck=false,QueryInstrStatus=true,Simulate=false")

            Dim result As Boolean
            result = driver.Initialized

            Console.Write("Done." & vbCrLf)


            outputPtr = driver.Outputs.Item(driver.Outputs.Name(channel))
            protectionPtr = driver.Protections.Item(driver.Protections.Name(channel))
            measurementPtr = driver.Measurements.Item(driver.Measurements.Name(channel))

            ' set voltage level
            outputPtr.VoltageLevel(3.0#, 3.0#)

            ' enable OV protection and set OV level
            protectionPtr.ConfigureOVP(10)

            'enable the over current protection
            protectionPtr.CurrentLimitBehavior = Agilent.AgilentN67xx.Interop.AgilentN67xxCurrentLimitBehaviorEnum.AgilentN67xxCurrentLimitTrip

            ' set current level
            outputPtr.CurrentLimit(1.0#, 1.0#)

            ' enable the output
            outputPtr.Enabled = True

            driver.Systems.WaitForOperationComplete(5000)

            ' Measure the voltage
            Dim measVoltage As Double
            measVoltage = measurementPtr.Measure(Agilent.AgilentN67xx.Interop.AgilentN67xxMeasurementTypeEnum.AgilentN67xxMeasurementVoltage)

            ' display the measured voltage
            Console.Write("Measured Voltage is " & measVoltage & " at channel " & channel & vbCrLf)

            'display instrument errors
            ReadInstrumentError(driver)

            'close the driver
            driver.Close()

            Console.Write("Driver closed." & vbCrLf)
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
            Console.Write(vbCrLf & "ErrorQuery: " & errCode & ", " & errMsg)
        End While

        Console.WriteLine()
    End Sub

End Module
