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

//  This program will only work with N676xA Power Modules and with N675xA Power Modules
//  that have Option 054 High Speed Test Extensions

//This program uses the voltage in step mode and also demonstrates 
//how to set up and use the digitizer.

#include <visa.h>
#include <stdio.h>
#include <stdlib.h>

int main()
{
	ViSession session,defrm;
	ViStatus stat;
	char statdsc[100];
	int channel,measpoints,measoffset;
	float voltage, finalvoltage;
	double timeinterval;
	char voltpoints[10000];

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
	
	//This controls the number of points the measurement system measures
	measpoints=100;

	//This controls the number of points to offset the measurement (positive for forward,
	//negative for reverse
	measoffset=0;

	//this sets the time between points
	timeinterval=0.0025;
	
	//this variable controls the voltage
	//it can be hard coded into the viPrintf command as well
	voltage=5;

	//This is the final voltage that will be triggered 
	finalvoltage=10;

	//this variable controls the channel
	//it can be hard coded into the viPrintf command as well
	channel=1;

	//The default resource manager manages all of the overhead required for VISA
	viOpenDefaultRM(&defrm);

	//opens a communication session with the instrument at address "add"
    stat=viOpen(defrm,IOaddress,VI_NULL,VI_NULL,&session);
	if (stat!=VI_SUCCESS) 
	{
		viStatusDesc(session,stat,statdsc);
		printf("Error on viOpen: %s \n",statdsc);
		exit(EXIT_FAILURE);
	}

	//Put the Voltage into step mode which causes it to transition from one
	//voltage to another upon receiving a trigger
	viPrintf(session,"VOLT:MODE STEP,(@%d) \n",channel);

	//program to voltage setting
	viPrintf(session,"VOLT %f,(@%d) \n",voltage,channel);

	//Go to final value
	viPrintf(session,"VOLT:TRIG %f,(@%d) \n",finalvoltage,channel);

	//Turn the output on
	viPrintf(session,"OUTP ON,(@%d) \n",channel);

	//Set the bus as the transient trigger source
	viPrintf(session,"TRIG:TRAN:SOUR BUS,(@%d) \n",channel);

	//Set the number of points for the measurement system to use as an offset
	viPrintf(session,"SENS:SWE:OFFS:POIN %d, (@%d) \n",measoffset,channel);

	//Set the number of points that the measurement system uses
	viPrintf(session,"SENS:SWE:POIN %d,(@%d) \n",measpoints,channel);

	//Set the time interval between points
	viPrintf(session,"SENS:SWE:TINT %f,(@%d) \n",timeinterval,channel);

	//Set the measurement trigger source
	viPrintf(session,"TRIG:ACQ:SOUR BUS,(@%d) \n",channel);
	
	//Initiate the measurement trigger system
	viPrintf(session,"INIT:ACQ (@%d) \n",channel);

	//Initiate the transient trigger system
	viPrintf(session,"INIT:TRAN (@%d) \n",channel);

	//Trigger the unit
	viPrintf(session,"*TRG \n");

	//Read back the voltage points
	viPrintf(session,"FETC:ARR:VOLT? (@%d) \n",channel);
	viScanf(session,"%s",&voltpoints);
	
	//Print the voltage points
	printf("There are the %d measured points: %s \n",measpoints,voltpoints);

	//free up all the resources
	viClose(session);
	viClose(defrm);

	return(EXIT_SUCCESS);
}
