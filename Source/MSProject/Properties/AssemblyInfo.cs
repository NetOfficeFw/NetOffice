using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	MSProject - Microsoft Project 11.0 Object Library - 11
	MSProject - Microsoft.Office.Interop.MSProject, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 12
	MSProject - Microsoft.Office.Interop.MSProject, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 14
	MSProject - Microsoft.Office.Interop.MSProject, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 14
	MSProject - Microsoft.Office.Interop.MSProject, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 16
*/

[assembly: AssemblyTitle("MSProject")]
[assembly: AssemblyDescription("NetOffice MSProject Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("MSProject")]
[assembly: Guid("A7107640-94DF-1068-855E-00DD01075445")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("VBIDEApi.dll", LoadHint.Default)]
[assembly: Dependency("MSHTMLApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/