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
'This program executes a 3 point current and voltage list.  It also specifies
'3 different dwell times.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

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
			Agilent.AgilentN67xx.Interop.IAgilentN67xxOutput3 outputPtr;
			Agilent.AgilentN67xx.Interop.IAgilentN67xxTransient  transientPtr;
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

				// get references to the needed interfaces
				outputPtr = driver.Outputs.get_Item(driver.Outputs.get_Name(channel))	;
				transientPtr = driver.Transients.get_Item(driver.Transients.get_Name(channel));

				//create arrays for the ListPoints method
				double[] voltList = new double [3];
				double[] currList = new double [3];
				double[] dwellTime = new double [3];

				voltList[0] = 5;
				voltList[1] = 7;
				voltList[2] = 10;

				currList[0] = 0.25;
				currList[1] = 0.5;
				currList[2] = 1.0;

				dwellTime[0] = 1;
				dwellTime[1] = 2;
				dwellTime[2] = 0.5;

				//' call ListPoints to set the list values and set the voltage and current modes to LIST
				transientPtr.ListPoints(ref voltList, ref currList, ref dwellTime);

				// enable the output
				outputPtr.Enabled = true;

				// Set the transient trigger source to bus.
				transientPtr.TrigSource = Agilent.AgilentN67xx.Interop.AgilentN67xxTriggerSourceEnum.AgilentN67xxTriggerSourceBus;

				// initiate the transient trigger system
				transientPtr.TrigInitiate();

				// trigger the transient system
				driver.Transients.SendSoftwareTrigger();

				driver.Systems.WaitForOperationComplete(5000);

				//Read insturment errors
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

			while (errCode != 0)
			{
				agDrvr.Utility.ErrorQuery(ref errCode, ref errMsg);
				Console.WriteLine("\n" +  "ErrorQuery: " + errCode + ", " + errMsg);
			}
		
		}
	}
}
