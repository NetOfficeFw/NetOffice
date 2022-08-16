﻿using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	DAO - Microsoft DAO 3.6 Object Library - 3.6
	DAO - <NoDescription> - 12.0

*/

[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("DAO")]
[assembly: Guid("00025E01-0000-0000-C000-000000000046")]
[assembly: NetOfficeAssemblyAttribute("1.8.1.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/