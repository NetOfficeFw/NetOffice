using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	ADODB - Microsoft ActiveX Data Objects 2.1 Library - 2.1
	ADODB - Microsoft ActiveX Data Objects 2.5 Library - 2.5

*/

[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("ADODB")]
[assembly: Guid("00000201-0000-0010-8000-00AA006D2EA4")]
[assembly: NetOfficeAssemblyAttribute("1.8.1.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/