using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Outlook - Microsoft Outlook 9.0 Object Library - 9
	Outlook - Microsoft Outlook 10.0 Object Library - 10
	Outlook - Microsoft.Office.Interop.Outlook, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	Outlook - Microsoft Outlook 12.0 Object Library - 12
	Outlook - Microsoft Outlook 14.0 Object Library - 14
	Outlook - Microsoft Outlook 15.0 Object Library - 15
    Outlook - Microsoft Outlook 15.0 Object Library - 16
*/

[assembly: AssemblyTitle("Outlook")]
[assembly: AssemblyDescription("NetOffice Outlook Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Outlook")]
[assembly: Guid("00062FFF-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/