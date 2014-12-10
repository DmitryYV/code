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
' This program demonstrates how to setup a transient and measurement trigger.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

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
			Agilent.AgilentN67xx.Interop.IAgilentN67xxTransient transientPtr;
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
				measurementPtr = driver.Measurements.get_Item(driver.Measurements.get_Name(channel));
				transientPtr = driver.Transients.get_Item(driver.Transients.get_Name(channel));

				//set the voltage level
				outputPtr.VoltageLevel(51, 5);

				//' Set the triggered voltage level and the mode
				//' Setting the voltage mode to step causes the voltage to transition from one
				//' voltage to another upon receiving a trigger
				transientPtr.set_StepTrigLevel (Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerLevelTypeEnum.AgilentN67xxTriggerLevelVoltage,10);

				//'enable the output
				outputPtr.Enabled = true;

				//' Set the transient trigger source to bus.
				transientPtr.TrigSource = Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerSourceEnum.AgilentN67xxTriggerSourceBus;

				//' initiate the transient trigger system
				transientPtr.TrigInitiate();

				//'Set trigger offset in the measurement sweep
				//'Set the number of data points in the measurement
				//'Set the measurement sample interval
				measurementPtr.Configure(Agilent.AgilentN67xx.Interop.AgilentN67xxMeasurementTypeEnum.AgilentN67xxMeasurementVoltage, 0, 5, 0.0025);

				//' set the measurement trigger source
				measurementPtr.TrigSource = Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerSourceEnum.AgilentN67xxTriggerSourceBus;

				//' initiate the measurement trigger system
				measurementPtr.TrigInitiate();

				//' trigger the measurement and transient systems
				driver.Transients.SendSoftwareTrigger();

				driver.Systems.WaitForOperationComplete(5000);

				//' read back the voltage points
				double[] fetchArray;
				fetchArray = measurementPtr.FetchArray(Agilent.AgilentN67xx.Interop.AgilentN67xxFetchTypeEnum.AgilentN67xxFetchVoltage);

				//' Print the voltage points
				string str="";
				int i;
				for (i = 0; i<= 4; i++)
				{
					str = str + fetchArray[i].ToString ();
					if (i != 4)
						str = str + ", ";
				}
				

				Console.Write("The measurement results are : " + str + "\n");

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
