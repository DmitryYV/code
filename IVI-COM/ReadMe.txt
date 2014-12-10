			***********************
			**** Read Me First ****
			***********************

Version: 1.1.0.19  March 28, 2007


Introducing the Agilent IVI-COM Driver for AgilentN67xx
-------------------------------------------------------
  The instrument driver provides access to the functionality of
  N67xx DC power supplies through a COM server which also complies with the IVI
  specifications. This driver works in any development environment which
  supports COM programming including Microsoft® Visual Basic, Microsoft
  Visual C++, Microsoft .NET, Agilent VEE Pro, National Instruments LabView,
  National Instruments LabWindows, and others.


Supported Instruments
---------------------
    Frame :
	N6700A
	N6700B
	N6701A
	N6702A 
	N6705A 

    Module :
	N6751A-N6752A
	N6761A-N6762A
	N6731A-N6735A
	N6742A-N6745A
	N6731B-N6736B
	N6741B-N6746B
	N6773A-N6776A 


Installation
-------------
  System Requirements: The driver installation will check for the
  following requirements.  If not found, the installer will either
  abort, warn, or install the required component as appropriate.

  Supported Operating Systems:
    Windows® 98 
    Windows Me
    Windows NT® 4.0 (SP6 or later)
    Windows 2000
    Windows XP

  Disc Space:
    IVI Shared Components: 8.4 Mb
    Driver:            2.68 MB
    The shared components are installed with the first driver
    installed.

  Shared Components
    Before this driver can be installed, your computer must already
    have several components already installed. The easiest way to
    to install these other components is by running the IVI Shared
    Components installation package available from the IVI Foundation
    as noted below. It installs all the items mentioned below.

    IVI Shared Components version 1.2.1.0 or later.
      The driver installer requires these components. They must be
      installed before the driver installer can install the driver.  

      The IVI Shared Components installer is available from: 
      http://www.ivifoundation.org/Downloads/SharedComponents.htm

      To install, open or run the file: IviSharedComponents.msi

      If the driver installer exhibits problems with installing the IVI
      Shared Components, find the program IviSharedComponents.msi and
      Run the program. It is able to install many of the required system
      components, including Windows Installer 2.0.

    Windows Installer
      Version 2.0 or later. The IVISharedComponents.exe program will
      install version 2.0 of the Windows installer.
      It is also available from:
       http://microsoft.com/downloads
      Enter "windows installer" as keywords.  Use the one appropriate for
      your Operating System.

    Microsoft XML 4.0
      The Configuration Server, an IVI Shared Component, uses Microsoft's
      XML 4.0 parser. The IVI Shared Components will install the XML 4.0
      parser. You can download the parser and its associated SDK from: 
       http://microsoft.com/downloads
      Enter "xml 4.0" as keywords.

    Microsoft HTML Help Viewer (Win98, WinNT) Version 1.32 or later
      The driver will install without the help viewer, but you will not
      be able to use the help file for the driver. You can download the
      help viewer from:
        http://microsoft.com/downloads
      Enter "html help workshop" as keywords

    MS Windows® Script
      Typically, wscript.exe is installed as part of the operating system
      or with Internet Explorer. If you see an error about wscript.exe,
      you need to install the Windows Script.
      Download from:
        http://msdn.microsoft.com/downloads
      Enter "windows script" as keywords. Use the one appropriate for
      your Operating System. 


  VISA-COM
    Any compliant implementation is acceptable. Typically, VISA-COM is
    installed with VISA and other I/O library modules. 

    If you are using Agilent Technologies' VISA-COM, use
    version 2.4.0.4 or later. The latest version of Agilent's IO
    Libraries is recommended and includes the required version of
    VISA-COM.

  I/O
    Agilent IO Libraries Suite, version 14.0 or later is recommended.
    You can download the latest version from:
      http://www.agilent.com/find/iolibsforivicom 

    If you are using National Instruments I/O libraries, use NI-VISA
    version 3.0.1 or later and NI-488.2 version 2.2 or later.

  Note: Driver may operate without any I/O software in simulation mode.

Additional Setup
----------------
  .NET Framework
    The .NET Framework itself is not required by this driver. If you
    plan to use the driver with .NET, Service Pack 2 is required.
    Download from:
 http://msdn.microsoft.com/netframework/downloads/updates/sp/default.asp

  The .NET Framework requires an interop assembly for a COM
  server. A Primary Interop Assembly, along with an XML file for
  IntelliSense is installed with the driver. The driver's PIA, along
  with IVI PIAs are installed, by default, in:
  <drive>:\Program Files\IVI\Bin\Primary Interop Assemblies

  The PIA is also installed into the Global Assembly Cache (GAC) if
  you have the .NET framework installed.

  If you install the .NET framework later, you can use the file:
  <drive>:\Program Files\IVI\Drivers\AgilentN67xx\DotNet.bat
  to put the driver's PIA into the GAC and properly register it.


Start Menu Help File 
--------------------
  A shortcut to the driver help file is added to the Start Menu,
  Programs, IVI, AgilentN67xx group.  It contains "Getting
  Started" information on using the driver in a variety of programming
  environments as well as documentation on IVI and instrument specific
  methods and properties.

  You will also see shortcuts to the Readme file and any programming
  examples for this driver.


MSI Installer
-------------
  The installation package for the driver is distributed as an MSI 2.0
  file. You can install the driver by double-clicking on the file.
  This operation actually runs:
      msiexec /i AgilentN67xx.msi

  You can run msiexec from a command prompt and utilize its many
  command line options. There are no public properties which can be
  set from the command line. You can also install this driver from
  another installation package.


Uninstall
---------
  This driver can be uninstalled like any other software from the
  Control Panel using "Add or Remove Programs". 

  The IVI Shared Components may also be uninstalled like any other
  software from the Control Panel using "Add or Remove Programs".

  Note: All IVI-COM drivers require the IVI Shared Components to
  function. To completely remove IVI components from your computer,
  uninstall all drivers and then uninstall the IVI Shared Components.


Using a New Version of the Driver
---------------------------------
  New versions of this Agilent IVI-COM driver may have a new
  ProgId.

  If you use the version dependent ProgId in CoCreateInstance,
  you will need to modify and recompile your code to use the
  new ProgID once you upgrade to the next version of the driver.
  Doing a side-by-side installation of the driver to use multiple
  versions of the driver is not supported.  If you need to go back
  to an older version of the driver, you need to uninstall the later
  version and install the older version.

  If you use the version independent ProgId in CoCreateInstance,
  you will not need to modify and recompile your code.  The new
  version of the driver has been tested to be backwards compatible
  with previous versions.

  To access the new functionality in a new version of the driver
  you will need to use the latest numbered IAgilentN67xx[n]
  interface rather than the IAgilentN67xx interface.  The
  IAgilentN67xx[n].<interface> property will return a pointer
  to the new IAgilentN67xx<interface>[n] interface.  The
  IAgilentN490x<interface>[n] interface contains the methods and 
  properties for the new functionality.  The new interfaces were
  introduced rather than modifying the existing interfaces for
  backwards compatibility.  The interfaces that were previously
  shipped have not been changed.


Known Issues
------------
  This driver does not have any known defects.


Revision History
----------------
  Version	Date		Notes
  -------	----		-----
  1.1.0.19	28 Mar 2007	Removed known defects	
  1.1.0.15	21 Mar 2007	Removed known defects	
  1.1.0.14	16 Mar 2007	Removed known defects
  1.1.0.13	15 Mar 2007	Removed Known Defects	
  1.1.0.12	07 Mar 2007	Removed Known Defects	
  1.1.0.10	07 Feb 2007	Added new Model Number N6705A	
  1.1.0.7	20 Dec 2005	Added support for frames N6701A, N6702A and modules N6773A-N6776A
  1.0.1.6	15 Nov 2004	Updated installer, added VISA-COM 3.0 PIA.
  1.0.1.5	10 Aug 2004	Added functionality for the N6700B frame.
				The tracing capability controlled by the systems interface
 				TraceEnabled property is not supported in this driver.
  1.0.1.0	12 Feb 2004	Bug fix for the Statuses Interface preset method.
  1.0.0.1       21 Jan 2004	The help file was updated, however, AgilentN67xx.dll remains
				unchanged and is version 1.0.0.0
  1.0.0.0	05 Dec 2003	Initial release


IVI Shared Component Revisions
------------------------------
  1.3.2.4   IviConfigServer.dll
  1.0.7.0   IviConfigServerCAPI.dll
  1.1.0.0   IviCShared.dll
  1.0.237.0 IviCSharedSupport.dll
  2.0.0.0   IviDCPwrTypeLib.dll
  3.0.0.0   IviDmmTypeLib.dll
  1.0.0.0   IviDriverTypeLib.dll
  1.1.1.0   IviEventServer.exe
  1.1.1.0   IviEventServerDLL.dll
  1.1.1.0   IviEventServerDLLps.dll
  1.1.1.0   IviEventServerps.dll
  3.0.0.0   IviFgenTypeLib.dll
  1.0.236.0 IviFloat.dll
  1.0.0.0   IviPwrMeterTypeLib.dll
  1.1.0.0   IviRfSigGenTypeLib.dll
  3.0.0.0   IviScopeTypeLib.dll
  1.0.1.2   IviSessionFactory.dll
  1.2.1.0   IviSharedComponentVersion.dll
  1.0.0.0   IviSpecAnTypeLib.dll
  3.0.0.0   IviSwtchTypeLib.dll


IVI Compliance
--------------
  (The following information is required by IVI 3.1 section 5.21.)

  IVI-COM Ivi DC Power Instrument Driver
  IVI Instrument Class: IVI-4.4_DCPwr_v2.0
  Group Capabilities:
  Supported:				
	IviDCPwrOutputs
	IviDCPwrOutput
	IviDCPwrTrigger

  Optional Features: This driver does not support Interchangeability
  Checking, State Caching, or Coercion Recording.

  Driver Identification: 
  (These three strings are values of properties in the IIviIdentity
  interface.)
  Vendor: Agilent Technologies. 
  Description: IVI-COM Driver for Agilent Technologies N67xx Power Supplies
	Revision: 1.1

  Component Identifier: AgilentN67xx.AgilentN67xx
  (The component identifier can be used to create an instance of the COM
  server.) 

  Hardware: This driver supports instruments manufactured by Agilent
  Technologies. The supported model numbers are:
  	N6700A
    	N6700B
	N6701A
	N6702A 

  This driver supports communicating with the instrument
  using GPIB or LAN.

  Software: See the section on installation in this document for
  information on what other software is required by this driver.  

  Source for this driver is not available.


More Information
----------------
  For more information about this driver and other instrument drivers
  and software available from Agilent Technologies visit:
  http://www.agilent.com/find/adnivicominfo

  A list of contact information is available from:
    http://www.agilent.com/find/assist

  Microsoft, Windows, MS Windows, and Windows NT are U.S. 
  registered trademarks of Microsoft Corporation.

© Copyright Agilent Technologies, Inc. 2004-2007