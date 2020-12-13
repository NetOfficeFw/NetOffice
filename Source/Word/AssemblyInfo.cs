﻿using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Word - Microsoft Word 9.0 Object Library - 9
	Word - Microsoft Word 10.0 Object Library - 10
	Word - Microsoft.Office.Interop.Word, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	Word - Microsoft Word 12.0 Object Library - 12
	Word - Microsoft Word 14.0 Object Library - 14
	Word - Microsoft Word 15.0 Object Library - 15
    Word - Microsoft Word 16.0 Object Library - 16
*/

[assembly: AssemblyTitle("Word")]
[assembly: AssemblyDescription("NetOffice Word Api")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("NetOfficeFw")]
[assembly: AssemblyProduct("NetOffice")]
[assembly: AssemblyCopyright("Copyright © 2012-2018 Sebastian Lange, © 2015-2020 Jozef Izso")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyVersion("1.8.0.0")]
[assembly: AssemblyFileVersion("1.8.0.0")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Word")]
[assembly: Guid("00020905-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.8.0.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("VBIDEApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/