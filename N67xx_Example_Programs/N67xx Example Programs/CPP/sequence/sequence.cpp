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

//This program sets up two output channels to turn on in sequence
//The first ouput channel turns on when the OUTP:ON command is received
//The second turns on 50 ms later.  All variables can be changed by the user

#include <visa.h>
#include <stdio.h>
#include <stdlib.h>

int main()
{
	//This block sets up all of the required variables
	ViSession session, defRM;
	ViStatus VISAstatus;
	int chA, chB;
	double voltA,voltB,currA,currB, ondelayB;
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

	//The two channels to be programmed are 1 and 2
	chA=1;
	chB=2;

	//Channel one at 5 V
	//Channel two at 6 V
	voltA=5;  //in volts
	voltB=6;  //in volts

	//Current on channel one is 1 A
	//Current on channel two is 2 A
	currA=1; //in amps
	currB=2; //in amps
     
	//There will be a 50 ms delay between the two channels 
	//turning on
	ondelayB=0.05; //in seconds

	//The following two commands set up VISA to communicate with the
	//instrument
	viOpenDefaultRM(&defRM);

	VISAstatus=viOpen(defRM,IOaddress,VI_NULL,VI_NULL,&session);
	if (VISAstatus!=VI_SUCCESS)
	{
		viStatusDesc(session,VISAstatus,statdsc);
		printf("Error on viOpen: %s \n",statdsc);
		exit(EXIT_FAILURE);
	}


	//Set voltage and current for both channels
	viPrintf(session,"VOLT %f,(@%d) \n",voltA,chA);
	viPrintf(session,"VOLT %f,(@%d) \n",voltB,chB);
	viPrintf(session,"CURR %f,(@%d) \n",currA,chA);
	viPrintf(session,"CURR %f,(@%d) \n",currB,chB);

	//This command sets the output on delay
	//An output off delay can also be set
	viPrintf(session,"OUTP:DEL:RISE %f,(@%d) \n",ondelayB,chB);

	//Turn the output on
	viPrintf(session,"OUTP ON,(@%d,%d) \n",chA,chB);

	//close down all of the VISA sessions
	viClose(session);
	viClose(defRM);

	return EXIT_SUCCESS;
}
	