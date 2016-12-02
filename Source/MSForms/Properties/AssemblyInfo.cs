using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	MSForms - Microsoft Forms 2.0 Object Library - 2

*/

[assembly: AssemblyTitle("MSForms")]
[assembly: AssemblyDescription("NetOffice MSForms Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("MSForms")]
[assembly: Guid("0D452EE1-E08F-101A-852E-02608C4D0BB4")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
[SupportByVersionAttribute("MSForms", 2)]
OLE_COLOR as Int32

[SupportByVersionAttribute("MSForms", 2)]
OLE_HANDLE as Int32

[SupportByVersionAttribute("MSForms", 2)]
OLE_OPTEXCLUSIVE as bool

[SupportByVersionAttribute("MSForms", 2)]
PIROWSET as object

*/