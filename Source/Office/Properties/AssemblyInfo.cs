using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Office - Microsoft Office 9.0 Object Library - 9
	Office - Microsoft Office 10.0 Object Library - 10
	Office - Microsoft Office 11.0 Object Library - 11
	Office - Office, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 12
	Office - Office, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 14
	Office - Microsoft Office 15.0 Object Library - 15
    Office - Microsoft Office 16.0 Object Library - 16
*/

[assembly: AssemblyTitle("Office")]
[assembly: AssemblyDescription("NetOffice Office Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Office")]
[assembly: Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
MsoRGBType as Int32

*/