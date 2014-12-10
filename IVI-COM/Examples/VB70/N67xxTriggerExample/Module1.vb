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
' This program demonstrates how to setup a transient and measurement trigger.
'/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Module Module1
    ' Agilent IVI-COM Driver Example Program
    ' 
    ' Creates a driver object and reads the Identifier and
    ' InstrumentFirmwareRevision properties.
    ' You must add a COM reference to the driver's type library.
    Private driver As Agilent.AgilentN67xx.Interop.AgilentN67xx
    Public outputPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxOutput3
    Public transientPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxTransient
    Public measurementPtr As Agilent.AgilentN67xx.Interop.IAgilentN67xxMeasurement

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
                                    "Cache=true,InterchangeCheck=false,QueryInstrStatus=true,Simulate=true")

            Dim result As Boolean
            result = driver.Initialized

            Console.Write("Done." & vbCrLf)


            outputPtr = driver.Outputs.Item(driver.Outputs.Name(channel))
            measurementPtr = driver.Measurements.Item(driver.Measurements.Name(channel))
            transientPtr = driver.Transients.Item(driver.Transients.Name(channel))

            'set the voltage level
            outputPtr.VoltageLevel(51, 5)

            ' Set the triggered voltage level and the mode
            ' Setting the voltage mode to step causes the voltage to transition from one
            ' voltage to another upon receiving a trigger
            transientPtr.StepTrigLevel(Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerLevelTypeEnum.AgilentN67xxTriggerLevelVoltage) = 10

            'enable the output
            outputPtr.Enabled = True

            ' Set the transient trigger source to bus.
            transientPtr.TrigSource = Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerSourceEnum.AgilentN67xxTriggerSourceBus

            ' initiate the transient trigger system
            transientPtr.TrigInitiate()

            'Set trigger offset in the measurement sweep
            'Set the number of data points in the measurement
            'Set the measurement sample interval
            measurementPtr.Configure(Agilent.AgilentN67xx.Interop.AgilentN67xxMeasurementTypeEnum.AgilentN67xxMeasurementVoltage, 0, 5, 0.0025)

            ' set the measurement trigger source
            measurementPtr.TrigSource = Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerSourceEnum.AgilentN67xxTriggerSourceBus

            ' initiate the measurement trigger system
            measurementPtr.TrigInitiate()



            ' trigger the measurement and transient systems
            driver.Transients.SendSoftwareTrigger()

            driver.Systems.WaitForOperationComplete(5000)

            ' read back the voltage points
            Dim fetchArray() As Double
            fetchArray = measurementPtr.FetchArray(Agilent.AgilentN67xx.Interop.AgilentN67xxFetchTypeEnum.AgilentN67xxFetchVoltage)

            ' Print the voltage points
            Dim str As String
            Dim i As Integer
            For i = 0 To 4
                str = str & fetchArray(i)
                If i <> 4 Then
                    str = str & ", "
                End If
            Next i

            Console.Write("The measurement results are : " & str & vbCrLf)


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
