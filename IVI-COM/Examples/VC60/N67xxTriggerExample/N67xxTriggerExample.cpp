/****************************************************************************
	'    Copyright © 2003-05 Agilent Technologies Inc. All rights
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
 ****************************************************************************/

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 This program demonstrates how to setup a transient and measurement trigger.
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

#include <stdafx.h>
#include <stdio.h>
#include <conio.h>
#include <atlbase.h>

#import "IviDriverTypeLib.dll" no_namespace
#import "IviDCPwrTypeLib.dll" no_namespace
#import "GlobMgr.dll" no_namespace
#import "AgilentN67xx.dll" no_namespace


int main()
{
	HRESULT hr;

	IAgilentN67xx3Ptr driverPtr=NULL;

	try {
		hr = CoInitialize(NULL);
		if (FAILED(hr))
			exit(1);

		// create a safe pointer for interface access
		hr = driverPtr.CreateInstance(__uuidof(AgilentN67xx));
		if (FAILED(hr))
			exit(1);

		// Get Identifier property.  Driver initialization not required.
		BSTR bstrIdn;
		bstrIdn = driverPtr->Identity->Identifier;
		wprintf(L"Identifier: %s\n", bstrIdn);

		// open the instrument for communication
		hr = driverPtr->Initialize(
			L"GPIB0::5::INSTR",  //Visa address,(not applicable if simulation=true)
			VARIANT_FALSE,    // ID query
			VARIANT_TRUE,     // Reset
			LPCTSTR("Cache=true, InterchangeCheck=false, QueryInstrStatus=true, Simulate=true")); //IVI options
		if (FAILED(hr))
			exit(1);

		// Get InstrumentFirmwareRevision property.
		BSTR bstrInstrFwRev;
		bstrInstrFwRev = driverPtr->Identity->InstrumentFirmwareRevision;
		wprintf(L"InstrumentFirmwareRevision: %s\n", bstrInstrFwRev);

		// get pointers to the needed interfaces	
		IAgilentN67xxOutput3Ptr outputPtr;
		outputPtr = driverPtr->Outputs->GetItem(OLESTR("Output1"));
		
		IAgilentN67xxTransientPtr transientPtr;
		transientPtr = driverPtr->Transients->GetItem(OLESTR("Transient1"));

		IAgilentN67xxMeasurementPtr measurementPtr;
		measurementPtr = driverPtr->Measurements->GetItem(OLESTR("Measurement1"));


		// set the voltage level
		hr = outputPtr->VoltageLevel(51.0, 5.0);
		if (FAILED(hr))
			exit(1);

		// set the triggered voltage level and the mode
		// Setting the voltage mode to step causes the voltage to transition from one
		// voltage to another upon receiving a trigger
		hr = transientPtr->put_StepTrigLevel(AgilentN67xxTriggerLevelVoltage, 10);

		// enable the output
		outputPtr->Enabled = VARIANT_TRUE;

		// Set the transient trigger source to bus.
		transientPtr->TrigSource = AgilentN67xxTriggerSourceBus;

		// initiate the transient trigger system
		hr = transientPtr->TrigInitiate();

		//Set trigger offset in the measurement sweep
		//Set the number of data points in the measurement
		//Set the measurement sample interval
		hr = measurementPtr->Configure(AgilentN67xxMeasurementVoltage,  0, 5, 0.0025);
	
		// set the measurement trigger source
		hr = measurementPtr->put_TrigSource(AgilentN67xxTriggerSourceBus);

		// initiate the measurement trigger system
		hr = measurementPtr->TrigInitiate();

		Sleep (1000);

		// trigger the measurement and transient systems
		hr = driverPtr->Measurements->SendSoftwareTrigger();

		 // Create a safearray
		SAFEARRAY *fetchArray = NULL;
		fetchArray = SafeArrayCreateVector(VT_R8, 0, 5);
		if (fetchArray == NULL)
			exit(1);

		// read back the voltage points
		fetchArray = measurementPtr->FetchArray(AgilentN67xxFetchVoltage);
		if (FAILED(hr))
			exit(1);

		// Get a pointer to the elements of the array.
		double HUGEP *pData;
		hr = SafeArrayAccessData(fetchArray, (void HUGEP* FAR*)&pData);
		if (FAILED(hr))
		   exit(1);

		// print out the array of data
		for (int i =0; i < 5; i++)
			printf("Element %d = %f \n", i+1, pData[i]);

		SafeArrayUnaccessData(fetchArray);

		::SafeArrayDestroy(fetchArray);

		//Read instrument errors
USES_CONVERSION;
		CComBSTR bstrError;
		long lErrCode=-1;

		while(lErrCode!=0)
		{
			driverPtr->Utility->ErrorQuery(&lErrCode,&bstrError);
			printf("\n%d, %s",lErrCode,OLE2A(bstrError));
		}

	
		hr = driverPtr->Close();
		if (FAILED(hr))
			exit(1);

		printf("\n\nPress any key to exit...");
		getch();

		CoUninitialize();
	
	}
	catch(_com_error err) {
		printf("%s", err.Description());
	}




	return 0;
}
