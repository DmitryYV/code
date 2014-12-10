@Rem This batch file removes the driver's Primary Interop Assembly, PIA,
@Rem from the Global Assembly Cache using gacutil.

@Rem Running this batch file makes the driver unusable with .NET.
@Rem This batch file reverses the operations done by the DotNet.bat file
@Rem found in the same directory.

gacutil -nologo -u Agilent.AgilentN67xx.Interop
