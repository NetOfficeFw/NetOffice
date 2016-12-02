using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	VBIDE - Microsoft Visual Basic for Applications Extensibility 5.3 - 5.3
	VBIDE - Microsoft.Vbe.Interop, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 12
	VBIDE - Microsoft.Vbe.Interop, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 14

*/

[assembly: AssemblyTitle("VBIDE")]
[assembly: AssemblyDescription("NetOffice VBIDE Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("VBIDE")]
[assembly: Guid("0002E157-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/