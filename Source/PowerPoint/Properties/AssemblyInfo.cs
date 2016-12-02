using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice;

/*
    Contains the following Type Libraries:
	Name - Description - SupportByVersion
	PowerPoint - Microsoft PowerPoint 9.0 Object Library - 9
	PowerPoint - Microsoft PowerPoint 10.0 Object Library - 10
	PowerPoint - Microsoft.Office.Interop.PowerPoint, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c - 11
	PowerPoint - Microsoft PowerPoint 12.0 Object Library - 12
	PowerPoint - <NoDescription> - 14
	PowerPoint - <NoDescription> - 15
    PowerPoint - <NoDescription> - 16
*/

[assembly: AssemblyTitle("PowerPoint")]
[assembly: AssemblyDescription("NetOffice PowerPoint Api")]
[assembly: PrimaryInteropAssembly(1, 0)]
[assembly: ImportedFromTypeLib("PowerPoint")]
[assembly: Guid("91493440-5A91-11CF-8700-00AA0060263B")]
[assembly: NetOfficeAssemblyAttribute("1.7.4.0")]
[assembly: Dependency("OfficeApi.dll", LoadHint.Default)]
[assembly: Dependency("VBIDEApi.dll", LoadHint.Default)]
[assembly: Dependency("NetOffice.dll", LoadHint.Default)]


/*
Alias Table
 
*/