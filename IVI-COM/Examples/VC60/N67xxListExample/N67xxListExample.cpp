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
This program executes a 3 point current and voltage list.  It also specifies
3 different dwell times.
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
	IAgilentN67xx3Ptr driverPtr=NULL;

	try 
	{
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
			LPCTSTR("Cache=true, InterchangeCheck=false, QueryInstrStatus=true, Simulate=false")); //IVI options
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

		// create safearrays to store the list data
		// voltage list
		SAFEARRAY *voltList = NULL;
		voltList = SafeArrayCreateVector(VT_R8, 0, 3);
		if (voltList == NULL)
			exit(1);

		double HUGEP* pd;
		hr = SafeArrayAccessData(voltList, (void HUGEP**)&pd);
		if (FAILED(hr)) {
			SafeArrayDestroy(voltList);
			exit(1);
		}

		*pd = 5;
		*(pd+1) = 7;
		*(pd+2) = 10;
		SafeArrayUnaccessData(voltList);


		// current list
		SAFEARRAY *currList = NULL;
		currList = SafeArrayCreateVector(VT_R8, 0, 3);
		if (currList == NULL) {
			 SafeArrayDestroy(voltList);
			 exit(1);
		}

		hr = SafeArrayAccessData(currList, (void HUGEP**)&pd);
		if (FAILED(hr)) {
			SafeArrayDestroy(currList);
			SafeArrayDestroy(voltList);
			exit(1);
		}

		*pd = 0.25;
		*(pd+1) = 0.50;
		*(pd+2) = 1.0;
		SafeArrayUnaccessData(currList);


		// dwell list
		SAFEARRAY *dwellTime = NULL;
		dwellTime = SafeArrayCreateVector(VT_R8, 0, 3);
		if (dwellTime == NULL) {
			 SafeArrayDestroy(voltList);
			 SafeArrayDestroy(currList);			 
			 exit(1);
		}

		hr = SafeArrayAccessData(dwellTime, (void HUGEP**)&pd);
		if (FAILED(hr)) {
			SafeArrayDestroy(dwellTime);
			SafeArrayDestroy(voltList);
			SafeArrayDestroy(currList);
			exit(1);
		}

		*pd = 1;
		*(pd+1) = 2;
		*(pd+2) = 0.5;
		SafeArrayUnaccessData(dwellTime);


		// call ListPoints to set the list values and set the voltage and current modes to LIST
		hr = transientPtr->ListPoints(&voltList, &currList, &dwellTime);
		if (FAILED(hr))
			exit(1);

		// enable the output
		outputPtr->Enabled = true;

		// Set the transient trigger source to bus.
		transientPtr->TrigSource = AgilentN67xxTriggerSourceBus;

		// initiate the transient trigger system
		hr = transientPtr->TrigInitiate();

		Sleep (1000);

		// trigger the transient system
		hr = driverPtr->Transients->SendSoftwareTrigger();

		// Destroy SAFEARRAYs
		::SafeArrayDestroy(voltList);
		::SafeArrayDestroy(dwellTime);
		::SafeArrayDestroy(currList);  

		//Read any instrument error
USES_CONVERSION;
		CComBSTR bstrError;
		long lErrCode=-1;

		while(lErrCode!=0)
		{
			driverPtr->Utility->ErrorQuery(&lErrCode,&bstrError);
			printf("\n%d, %s",lErrCode,OLE2A(bstrError));
		}

		printf("\n\nPress any key to exit...");
		getch();


	}
	catch(_com_error err) 
	{
		printf("%s", err.Description());
	}

	hr = driverPtr->Close();
	if (FAILED(hr))
		exit(1);

	CoUninitialize();
	return 0;
}