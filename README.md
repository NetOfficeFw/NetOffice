# NetOffice - Microsoft Office in .NET

[![Build Status](https://dev.azure.com/netoffice/NetOffice/_apis/build/status/NetOfficeFw.NetOffice?branchName=master)](https://dev.azure.com/netoffice/NetOffice/_build/latest?definitionId=1&branchName=master)

> .NET wrapper assemblies for accessing Microsoft Office applications.

NetOffice supports extending and automating Microsoft Office applications: Excel, Word, Outlook,
PowerPoint, Access, Project and Visio.

## Features

* MS Office integration without version limitations
* All features of the MS Office versions 2000, 2002, 2003, 2007, 2010, 2013 and 2016 are included
* Active support in version independent development
* Syntactically and semantically identical to the Microsoft Interop Assemblies
* No training if you already know the MS Office object model, use your existing PIA code
* Reduced and more readable code with automatic management of COM proxies
* Usable with .NET Framework 4.0 or higher
* Easy add-ins development
* No deployment hurdles, no registration
* No dependencies, no interop assemblies, no need for [VSTO][VSTO]
* Visual Studio Project Templates and Wizards available in [NetOffice Toolbox][NetOffice Toolbox]

## Getting Started

Checkout the [NetOffice-Examples](https://github.com/NetOfficeFw/NetOffice-Examples) repository
to see how to use NetOffice to automate Office applications or how to create add-ins to extend them.

## Tools

The [NetOffice Toolbox](https://github.com/NetOfficeFw/NetOfficeToolbox) is a comprehensive
toolset to get started with NetOffice solutions.

## Project History

You can find more information about [NetOffice Git repository in documentation](Documentation/History.md).

### Branches

* `master` - main branch
* `releases/netoffice_v1.7.4` - branch with current NetOffice 1.7.4 source code

#### Archive Branches

These branches are archived of the source code from CodePlex and OSDN.

* `import/osdn_repository` - branch with NetOffice source code imported from OSDN repository
* `import/legacy_repository` - archive branch of original NetOffice source code imported from CodePlex Subversion
* `import/netoffice_1.7.4-alpha` - archive branch of NetOffice 1.7.4 source code provided by Sebastian
* `import/netoffice_1.7.4.1-alpha` - archive branch of NetOffice 1.7.4.1 source code provided by Sebastian

## License

NetOffice source code is licensed under [MIT License](LICENSE.txt).

Copyright © 2011-2018 Sebastian Lange  
Copyright © 2015-2020 Jozef Izso


[VSTO]: https://docs.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio
[NetOffice Toolbox]: https://netoffice.io/toolbox/
