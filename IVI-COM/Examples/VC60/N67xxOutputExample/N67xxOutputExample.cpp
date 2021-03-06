/****************************************************************************
	'    Copyright � 2003-05 Agilent Technologies Inc. All rights
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
 ***************************************************************************/

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 This program demonstrates how to communicate with the System DC Source 
 using the IVI-COM driver. 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

#include <stdafx.h>
#include <stdio.h>
#include <conio.h>
#include <atlbase.h>

#import "IviDriverTypeLib.dll" no_namespace
#import "IviDCPwrTypeLib.dll" no_namespace
#import "GlobMgr.dll" no_namespace
#import "AgilentN67xx.dll" no_namespace


int main(int argc, char* argv[])
{
	HRESULT hr;

	try 
	{
		hr = CoInitialize(NULL);
		if (FAILED(hr))
			exit(1);

		// create a safe pointer for interface access
		IAgilentN67xx3Ptr driverPtr=NULL;
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

		IAgilentN67xxProtection2Ptr protectionPtr;
		protectionPtr = driverPtr->Protections->GetItem(OLESTR("Protection1"));

		IAgilentN67xxMeasurementPtr measurementPtr;
		measurementPtr = driverPtr->Measurements->GetItem(OLESTR("Measurement1"));
		
		// set voltage
		hr = outputPtr->VoltageLevel(3.0, 3.0);
		if (FAILED(hr))
			exit(1);

		// enable OV protection and set the limit.
		hr = protectionPtr->ConfigureOVP(10);
		if (FAILED(hr))
			exit(1);

		// enable the over current protection
		protectionPtr->CurrentLimitBehavior  = AgilentN67xxCurrentLimitTrip;

		// set the current limit
		hr = outputPtr->CurrentLimit(1.0, 1.0);
		if (FAILED(hr))
			exit(1);

		// enable the output
		outputPtr->Enabled = true;
		
		Sleep(100);

		// measure the voltage
		double measVoltage;
		measVoltage = measurementPtr->Measure(AgilentN67xxMeasurementVoltage);

		//print out voltage measurement
		printf("measured voltage : %f V \n",measVoltage);

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
	catch(_com_error err) 
	{
		printf("%s", err.Description());
	}

		
	return 0;
}