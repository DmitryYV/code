// """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
//  © Agilent Technologies, Inc. 2007

// You have a royalty-free right to use, modify, reproduce and distribute
// the Sample Application Files (and/or any modified version) in any way
// you find useful, provided that you agree that Agilent Technologies has no
// warranty,  obligations or liability for any Sample Application Files.
//
// Agilent Technologies provides programming examples for illustration only,
// This sample program assumes that you are familiar with the programming
// language being demonstrated and the tools used to create and debug
// procedures. Agilent Technologies support engineers can help explain the
// functionality of Agilent Technologies software components and associated
// commands, but they will not modify these samples to provide added
// functionality or construct procedures to meet your specific needs.
// """""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

//This is a simple program that sets a voltage, current, overvoltage, and the 
//status of over current protection.  All of the variables 
//can be changed by the user.

#include <visa.h>
#include <stdio.h>
#include <stdlib.h>

int main()
{
	ViSession session,defrm;
	ViStatus VISAstatus;
	char statdsc[100];
	char voltmeasurement[16]="0";
	int channel,overcurron,overvolton;
	float voltsetting,currsetting,overvoltsetting;
	
	//This variable is to be used for GPIB communication
	//GPIB0 is the VISA name of the interface 
	//The instrument is at address 5

	char IOaddress[]="GPIB0::5";

	//This variable can be changed for TCPIP communication
	//"TCPIP0::141.25.36.214" would be the string for TCPIP
	//TCPIP0 is the VISA name of the interface
	//141.25.36.214 is the IP address of the instrument

	//This variable can also be changed for USB communication
	//"USB0::2391::1799::MY00000002" would be the string for USB
	//USB0 is the VISA name of the interface
	//2391 is a unique identifier for Agilent
	//1799 is the instrument identifier
	//MY00000002 is the instrument serial number
	
	//this variable controls the voltage
	//it can be hard coded into the viPrintf command as well
	voltsetting=3; //in volts

	//this variable controls the channel
	//it can be hard coded into the viPrintf command as well
	channel=1;

	//this variable controls the current
	//it can be hard coded into the viPrintf command as well
	currsetting=0.5; //in amps
	
	//these vaiables control the over voltage protection setting
	overvolton=1;  //1 for on or 0 for off
	overvoltsetting=5;  //in volts

	//this variable enables/disables over-current protection 
	overcurron=0;	//1 for on, 0 for off

	//The default resource manager manages initializes the VISA system
	VISAstatus=viOpenDefaultRM(&defrm);

	//opens a communication session with the instrument at address "add"
    VISAstatus=viOpen(defrm,IOaddress,VI_NULL,VI_NULL,&session);
	if (VISAstatus!=VI_SUCCESS) 
	{
		viStatusDesc(session,VISAstatus,statdsc);
		printf("Error on viOpen: %s \n",statdsc);
		exit(EXIT_FAILURE);
	}

	//Set voltage
	viPrintf(session,"VOLT %f,(@%d) \n",voltsetting,channel);
	
	//Set the over voltage level
	viPrintf(session,"VOLT:PROT:LEV %f,(@%d) \n",overvoltsetting,channel);

	//Turn on over current protection 
	viPrintf(session,"CURR:PROT:STAT %d,(@%d) \n",overcurron,channel);
	
	//Set current level
	viPrintf(session,"CURR %f,(@%d) \n",currsetting,channel);
	
	//Enable the output
	viPrintf(session,"OUTP ON,(@%d) \n",channel);
	
	//measure the voltage
	viPrintf(session, "MEAS:VOLT? (@%d) \n",channel);
	

	//readback the voltage from the instrument
	viScanf(session,"%s",&voltmeasurement);

	//print out voltage measurement
	printf("voltage measured: %s \n",voltmeasurement);

	//This frees up all of the resources
	viClose(session);
	viClose(defrm);

	return EXIT_SUCCESS;
	

}