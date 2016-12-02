using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	MSComctlLib - Microsoft Windows Common Controls 6.0 - 6

*/

[assembly: AssemblyTitle("MSComctlLib")]
[assembly: AssemblyDescription("NetOffice MSComctlLib Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("MSComctlLib")]
[assembly: Guid("831FDD16-0C5C-11D2-A9FC-0000F8754DA1")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/