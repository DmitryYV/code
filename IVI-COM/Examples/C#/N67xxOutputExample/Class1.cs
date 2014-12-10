/*'/****************************************************************************
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
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' This program demonstrates how to communicate with a N67xx Modular Power System
' using the IVI-COM driver.
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

using System;
using Agilent.AgilentN67xx.Interop;

namespace N67xxListExample
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class Class1
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			// Create driver instance.
			AgilentN67xx driver;
			Agilent.AgilentN67xx.Interop.IAgilentN67xxProtection2 protectionPtr;
			Agilent.AgilentN67xx.Interop.AgilentN67xxMeasurement measurementPtr;
			Agilent.AgilentN67xx.Interop.IAgilentN67xxOutput3 outputPtr;
			
			driver = new AgilentN67xxClass();

			 
			int channel;  //' the channel to be programmed
			channel = 1;

			try
			{
				//' Get Identifier property.  Driver initialization not required.
				System.Console.Write("Identifier: " + driver.Identity.Identifier + "\n");
				System.Console.Write("Revision: " + driver.Identity.Revision + "\n");
				System.Console.Write("Description: " + driver.Identity.Description + "\n");
				System.Console.Write("Initializing... ");

				// Initialize driver 
				driver.Initialize("GPIB0::5::INSTR",
					false,
					true,
					"Cache=true,InterchangeCheck=false,QueryInstrStatus=true,Simulate=false");//Optional ivi options

				System.Console.WriteLine("Done.\n");

				//' Get InstrumentModel and InstrumentFirmwareRevision property.
				System.Console.Write("InstrumentModel:  " + driver.Identity.InstrumentModel + "\n");
				System.Console.Write("FirmwareRevision: " + driver.Identity.InstrumentFirmwareRevision + "\n");

				outputPtr = driver.Outputs.get_Item(driver.Outputs.get_Name(channel));
				protectionPtr = driver.Protections.get_Item(driver.Protections.get_Name(channel));
				measurementPtr = driver.Measurements.get_Item(driver.Measurements.get_Name(channel));

				//' set voltage level
				outputPtr.VoltageLevel(3.0, 3.0);

				//' enable OV protection and set OV level
				protectionPtr.ConfigureOVP(10);

				//'enable the over current protection
				protectionPtr.CurrentLimitBehavior = Agilent.AgilentN67xx.Interop.AgilentN67xxCurrentLimitBehaviorEnum.AgilentN67xxCurrentLimitTrip;

				//' set current level
				outputPtr.CurrentLimit(1.0, 1.0);

				//' enable the output
				outputPtr.Enabled = true;

				driver.Systems.WaitForOperationComplete(5000);

				//' Measure the voltage
				double measVoltage;
				measVoltage = measurementPtr.Measure(Agilent.AgilentN67xx.Interop.AgilentN67xxMeasurementTypeEnum.AgilentN67xxMeasurementVoltage);

	            //' display the measured voltage
				Console.Write("\nMeasured Voltage is " + measVoltage + " at channel " + channel );

				// Read insturment errors
				ReadInstrumentError(driver);

				//' Close driver if initialized.
				if (driver.Initialized == true ) driver.Close();

				Console.WriteLine();
				Console.Write("Press Enter to Exit ");
				Console.ReadLine();
			}
			catch (Exception e)
			{
				Console.WriteLine();
				Console.WriteLine("Exception Error:");
				Console.WriteLine("  " + e.Message );

				driver.Close();

				Console.WriteLine();
				Console.Write("Press Enter to Exit ");
				Console.ReadLine();
			}
		}

		static void ReadInstrumentError(Agilent.AgilentN67xx.Interop.AgilentN67xx agDrvr)
		{
			// Read instrument error queue until its empty.
			int errCode;
			errCode = 999;
			string errMsg="";
			Console.WriteLine();

			while (errCode != 0)
			{
				agDrvr.Utility.ErrorQuery(ref errCode, ref errMsg);
				Console.WriteLine("\n" +  "ErrorQuery: " + errCode + ", " + errMsg);
			}
		}
	}
}
