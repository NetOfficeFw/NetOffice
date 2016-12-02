using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Excel - Microsoft Excel 9.0 Object Library - 9
	Excel - Microsoft Excel 10.0 Object Library - 10
	Excel - Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	Excel - Microsoft Excel 12.0 Object Library - 12
	Excel - Microsoft Excel 14.0 Object Library - 14
	Excel - Microsoft Excel 15.0 Object Library - 15
    Excel - Microsoft Excel 16.0 Object Library - 15
*/

[assembly: AssemblyTitle("Excel")]
[assembly: AssemblyDescription("NetOffice Excel Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Excel")]
[assembly: Guid("00020813-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("VBIDEApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/