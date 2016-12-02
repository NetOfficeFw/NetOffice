using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Visio - <NoDescription> - 11
	Visio - <NoDescription> - 12
	Visio - <NoDescription> - 14
	Visio - <NoDescription> - 15
    Visio - <NoDescription> - 16
*/

[assembly: AssemblyTitle("Visio")]
[assembly: AssemblyDescription("NetOffice Visio Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Visio")]
[assembly: Guid("00021A98-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/