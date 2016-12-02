using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	OWC10 - Microsoft Office XP Web Components - 1

*/

[assembly: AssemblyTitle("OWC10")]
[assembly: AssemblyDescription("NetOffice OWC10 Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("OWC10")]
[assembly: Guid("0002E550-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("MSDATASRCApi.dll", LoadHint.Default)]
[assembly: Dependency("MSComctlLibApi.dll", LoadHint.Default)]
[assembly: Dependency("ADODBApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/