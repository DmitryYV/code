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
//  that have Option 054 High Speed Test Extensions and are rated for the specified
//  voltage and current combinations embedded in the code.

//This program executes a 10 point current and voltage list.  It also specifies
//10 different dwell times. 

#include <visa.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>


int main()
{
	// These variable are neccessary to initialize the VISA system.
	ViStatus VISAstatus;
	ViSession session,defrm;
	
	//This variable controls the channel nuumber to be programmed
	int channel;

	//These next three strings are the points in the list.
	//All three strings are the same length.
	//The first one controls voltage, the second current, and the third dwell time
	char voltpts[]="1,2,3,4,5,6,7,8,9,10";
	char currpts[]="0.5,1,1.5,2,2.5,3,3.5,4,4.5,5";
	char dwellpts[]="1,2,0.5,1,0.25,1.5,0.1,1,0.75,1.2";

	char statdsc[100];

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

	//This variable can be changed to program any channel in the mainframe
	channel=3;

	//This function initializes the VISA system
	viOpenDefaultRM(&defrm);
	
	//Open a VISA session with the instrument
	VISAstatus=viOpen(defrm,IOaddress,VI_NULL,VI_NULL,&session);
	if (VISAstatus!=VI_SUCCESS)
	{
		viStatusDesc(session,VISAstatus,statdsc);
		printf("Error on viOpen: %s \n",statdsc);
		exit(EXIT_FAILURE);
	}
	
	//Set the voltage mode to list
	viPrintf(session,"VOLT:MODE LIST,(@%d) \n",channel);
	
	//Set the current mode to list
	viPrintf(session,"CURR:MODE LIST,(@%d) \n",channel);
	
	//Send the voltage list points
	viPrintf(session,"LIST:VOLT %s,(@%d) \n",voltpts,channel);
	
	//Send the Current list points
	viPrintf(session,"LIST:CURR %s,(@%d) \n",currpts,channel);
	
	//Send the dwell points
	viPrintf(session,"LIST:DWEL %s,(@%d) \n",dwellpts,channel);
	
	//Turn the output on
	viPrintf(session,"OUTP ON,(@%d) \n",channel);
	
	//Set the trigger source to bus
	viPrintf(session,"TRIG:TRAN:SOUR BUS,(@%d) \n",channel);
	
	//Initiate the transoent system
	viPrintf(session,"INIT:TRAN (@%d) \n",channel);
	
	//Trigger the unit
	viPrintf(session,"*TRG \n");
	
	//Free up all of the resources that were assigned to the VISA system.
	viClose(session);
	viClose(defrm);

	return EXIT_SUCCESS;
}