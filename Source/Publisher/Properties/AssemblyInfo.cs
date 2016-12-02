using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Publisher - Microsoft.Office.Interop.Publisher, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 14
	Publisher - Microsoft Publisher 15.0 Object Library - 15
	Publisher - Microsoft Publisher 16.0 Object Library - 16
*/

[assembly: AssemblyTitle("Publisher")]
[assembly: AssemblyDescription("NetOffice Publisher Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Publisher")]
[assembly: Guid("0002123C-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/