# Changelog

## v1.7.5
* Fixed the `CoClassEventReflector` class implementation
* `CoClassEventReflector` class will throw `ArgumentOutOfRangeException` when event does not exist in the class (breaking-change)
* You can use the `nameof()` operator with event name in the `CoClassEventReflector` methods

## v1.7.4.11
* Change `TaskPaneInfo.TagHwnd` to `IntPtr` type

## v1.7.4.10
* Use portable symbol files in Release builds

## v1.7.4.9
* Allow `TaskPaneInfo` objects to be tagged from user code

## v1.7.4.6
* Fix #231 - Access library ProjectInfo returns incorrect AssemblyName value

## v1.7.4.5
* Fix #223 - OlRibbonType.cs wrong enum for Microsoft.Outlook.Mail.Compose

## v1.7.4.4
* MS Publisher package contains correct assemblies #216 (Publisher NuGet Package has WordApi.dll/xml and not PublisherApi.dll/xml)

## v1.7.4.2
* `COMAddin` supports custom addin object - see Word addin example **06 Custom Addin Object**

## v1.7.4.1
* Tutorials demonstrate most of the new core features (dynamics, cloning, etc)
* Skip support for old .NET Runtime versions - minimum is .NET 4.0 (Client Profile)  
  > We want to support .NET 4 (and any higher of course) as long as possible because it is the last Windows XP compatible runtime.
  > (NetOffice 1.7.3 with .NET 2.0/3.x support is still available in the [download section](https://github.com/NetOfficeFw/NetOffice/releases/tag/v1.7.3))
* Microsoft Publisher is now into play.
* Add **MSFormsApi.dll** to support VBE UI controls
* Total size of the assemblies is 25% smaller
* Extended support for MS-Excel RTD Server (see COMAddin examples)
* Extended support for _Document Inspector_ in MS-Word (see COMAddin examples)
* Extended support for custom MS Outlook property pages and Form Regions (see COMAddin examples)
* `[CustomUI]` attribute can handle Ribbon IDs now
* Suppress MS Outlook Security dialog is now available in `NetOffice.OutlookApi.Tools.Contribution.Security`
* Spend Contribution utils as optional service for common tasks
* **Developer Toolbox** source is available on <https://osdn.net/projects/netoffice> or SVN: <https://svn.osdn.net/svnroot/netoffice>
* Official mirror on GitHub is "netofficefw" - NOT "netoffice"
