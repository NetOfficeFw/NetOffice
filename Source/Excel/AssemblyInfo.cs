﻿using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	Excel - Microsoft Excel 9.0 Object Library - 9
	Excel - Microsoft Excel 10.0 Object Library - 10
	Excel - Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	Excel - Microsoft Excel 12.0 Object Library - 12
	Excel - Microsoft Excel 14.0 Object Library - 14
	Excel - Microsoft Excel 15.0 Object Library - 15
    Excel - Microsoft Excel 16.0 Object Library - 16
*/

[assembly: AssemblyTitle("Excel")]
[assembly: AssemblyDescription("NetOffice Excel Api")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("NetOfficeFw")]
[assembly: AssemblyProduct("NetOffice")]
[assembly: AssemblyCopyright("Copyright © 2012-2018 Sebastian Lange, © 2015-2020 Jozef Izso")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyVersion("1.7.5.1")]
[assembly: AssemblyFileVersion("1.7.5.1")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("Excel")]
[assembly: Guid("00020813-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.7.5.1")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("VBIDEApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/