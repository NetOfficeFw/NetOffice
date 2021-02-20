﻿using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	MSForms - Microsoft Forms 2.0 Object Library - 2

*/

[assembly: AssemblyTitle("MSForms")]
[assembly: AssemblyDescription("NetOffice MSForms Api")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("NetOfficeFw")]
[assembly: AssemblyProduct("NetOffice")]
[assembly: AssemblyCopyright("Copyright © 2012-2018 Sebastian Lange, © 2015-2021 Jozef Izso")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyVersion("1.8.0.0")]
[assembly: AssemblyFileVersion("1.8.0.0")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("MSForms")]
[assembly: Guid("0D452EE1-E08F-101A-852E-02608C4D0BB4")]
[assembly: NetOfficeAssemblyAttribute("1.8.0.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
[SupportByVersion("MSForms", 2)]
OLE_COLOR as Int32

[SupportByVersion("MSForms", 2)]
OLE_HANDLE as Int32

[SupportByVersion("MSForms", 2)]
OLE_OPTEXCLUSIVE as bool

[SupportByVersion("MSForms", 2)]
PIROWSET as object

*/