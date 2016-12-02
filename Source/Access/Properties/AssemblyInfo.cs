using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Access - Microsoft Access 9.0 Object Library - 9
	Access - Microsoft Access 10.0 Object Library - 10
	Access - Microsoft.Office.Interop.Access, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	Access - Microsoft Access 12.0 Object Library - 12
	Access - Microsoft Access 14.0 Object Library - 14
	Access - Microsoft Access 15.0 Object Library - 15
    Access - Microsoft Access 15.0 Object Library - 16
*/

[assembly: AssemblyTitle("Access")]
[assembly: AssemblyDescription("NetOffice Access Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Access")]
[assembly: Guid("4AFFC9A0-5F99-101B-AF4E-00AA003F0F07")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("DAOApi.dll", LoadHint.Default)]
[assembly: Dependency("VBIDEApi.dll", LoadHint.Default)]
[assembly: Dependency("ADODBApi.dll", LoadHint.Default)]
[assembly: Dependency("OWC10Api.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/