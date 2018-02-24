using System.Reflection;
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
[assembly: AssemblyDescription("Netoffice MSForms Api")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("https://osdn.net/projects/netoffice")]
[assembly: AssemblyProduct("NetOffice")]
[assembly: AssemblyCopyright("Sebastian Lange")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyVersion("1.7.4.4")]
[assembly: AssemblyFileVersion("1.7.4.4")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("MSForms")]
[assembly: Guid("0D452EE1-E08F-101A-852E-02608C4D0BB4")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.1")]
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