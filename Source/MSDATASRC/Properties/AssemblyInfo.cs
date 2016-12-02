using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	MSDATASRC - Microsoft Data Source Interfaces - 4

*/

[assembly: AssemblyTitle("MSDATASRC")]
[assembly: AssemblyDescription("NetOffice MSDATASRC Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("MSDATASRC")]
[assembly: Guid("7C0FFAB0-CD84-11D0-949A-00A0C91110ED")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
[SupportByVersionAttribute("MSDATASRC", 4)]
DataMember as string

*/