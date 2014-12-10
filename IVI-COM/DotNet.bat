@Rem This batch file puts the driver's Primary Interop Assembly, PIA, 
@Rem into the Global Assembly Cache using gacutil which is done by the
@Rem installer if the .NET Framework is already installed when the driver
@Rem is installed.

@Rem Running this batch file only makes sense if the .NET Framework is installed
@Rem after the driver is installed.

@Rem To undo the operations performed by this batch file, execute the
@Rem UnDotNet.bat file found in the same directory.

gacutil -nologo -i "C:\Program Files\IVI Foundation\IVI\Bin\Primary Interop Assemblies\Agilent.Itl.Interop.dll"
gacutil -nologo -i "C:\Program Files\IVI Foundation\IVI\Bin\Primary Interop Assemblies\Agilent.AgilentN67xx.Interop.dll"
