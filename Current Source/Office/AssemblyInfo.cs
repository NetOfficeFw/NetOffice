using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using LateBindingApi.Core;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Office - Microsoft Office 9.0 Object Library - 9
	Office - Microsoft Office 10.0 Object Library - 10
	Office - Office, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	Office - Office, Version=12.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C - 12
	Office - Microsoft Office 14.0 Object Library - 14

*/

[assembly: AssemblyTitle("Office")]
[assembly: AssemblyDescription("NetOffice Office Api")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("netoffice.codeplex.com")]
[assembly: AssemblyProduct("NetOffice")]
[assembly: AssemblyCopyright("Sebastian Lange")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyVersion("1.5.0.0")]
[assembly: AssemblyFileVersion("1.5.0.0")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Office")]
[assembly: Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52")]
[assembly: LateBindingAttribute("1.0")]

/*
Alias Table
 
[SupportByVersionAttribute("Office", 9,10,11,12,14)]
MsoRGBType as Int32

*/