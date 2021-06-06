# Changelog

## v2.0.0

* Enable support for compilation with .NET Core 3.1 [#219](https://github.com/NetOfficeFw/NetOffice/issues/219)

## v1.8.1

### Added
* Parser for the resiliency binary data stored by Microsoft Office when add-ins crash.

## v1.8.0

### Breaking Changes
* `Settings.EnableOperatorOverlads` was renamed to `Settings.EnableOperatorOverloads` [#306](https://github.com/NetOfficeFw/NetOffice/issues/306)  
  If you listened to changes using `Settings.PropertyChanged` event, ensure you fix your code with the new property name.
* Fixed typo in word **Availability** used in namespace and class names [#307](https://github.com/NetOfficeFw/NetOffice/issues/307)  
  Affected types:
  * namespace `NetOffice.Availity` 🡒 `NetOffice.Availability`
  * class `AvailityException` 🡒 `AvailabilityException`
  * interface `ICOMObjectAvaility` 🡒 `ICOMObjectAvailability`

## v1.7.9
* Links now point to Microsoft Docs website
* Documented PowerPoint types related to animations and effects
* Documented `LoadBehavior` values

## v1.7.8

* `Settings.IsEqualTo()` method will compare objects correctly [#291](https://github.com/NetOfficeFw/NetOffice/issues/291)
* `Settings` implements `IEquatable<Settings>` in favor of old `IsEqualTo()` method
* `ResourceUtils.ReadImage()` will correctly read the image from resources [#292](https://github.com/NetOfficeFw/NetOffice/issues/292)
* `ResourceUtils.ReadString()` will correctly read string resources [#292](https://github.com/NetOfficeFw/NetOffice/issues/292)

### Bug Fixes
* Fix [#291](https://github.com/NetOfficeFw/NetOffice/issues/291) _`Settings.IsEqualTo` compares `EnableSafeMode` incorrectly_
* Fix [#292](https://github.com/NetOfficeFw/NetOffice/issues/292) _`ResourceUtils.ReadImage()` calls itself recursively_

## v1.7.7

* Make the `CustomTaskPaneCollection.Remove()` method public to allow custom implementations of Task Panes to remove old objects manually

## v1.7.6

* Improved documentation for many members in the **NetOffice** project
* Refactored small portion of the `CurrentAppDomain.CurrentDomain_AssemblyResolve()` method
* Deprecated code related to dynamic types and "duck tales" implementation (will be removed in **v2.0**) [#283](https://github.com/NetOfficeFw/NetOffice/issues/283)

### Breaking Changes
* Some exceptions have better error messages and use correct parameters  
  _(if you rely on exact exception messages, this is a breaking change)_

## v1.7.5

### Added
* You can use the `nameof()` operator with event name in the `CoClassEventReflector` methods

### Breaking Changes
* `CoClassEventReflector` class will throw `ArgumentOutOfRangeException` when event does not exist in the class

### Bug Fixes
* Fix [#277](https://github.com/NetOfficeFw/NetOffice/issues/277) _CoClassEventReflector.HasEventRecipients always return false_

## v1.7.4.11
* Change `TaskPaneInfo.TagHwnd` to `IntPtr` type

## v1.7.4.10
* Use portable symbol files in Release builds

## v1.7.4.9
* Allow `TaskPaneInfo` objects to be tagged from user code

## v1.7.4.8

### Bug Fixes
* Fix [#262](https://github.com/NetOfficeFw/NetOffice/issues/262) _ActivePowerPointApp.SlideShowBeginEvent doesn't work in versions 1.7.4.x_

## v1.7.4.6

### Bug Fixes
* Fix [#231](https://github.com/NetOfficeFw/NetOffice/issues/231) - Access library ProjectInfo returns incorrect AssemblyName value
* Fix [#264](https://github.com/NetOfficeFw/NetOffice/issues/264) NetOfficeException "Keytoken missmatch" when running NetOffice in debug mode

## v1.7.4.5

### Bug Fixes
* Fix [#223](https://github.com/NetOfficeFw/NetOffice/issues/223) - _OlRibbonType.cs wrong enum for Microsoft.Outlook.Mail.Compose_

## v1.7.4.4
* MS Publisher package contains correct assemblies #216 (Publisher NuGet Package has WordApi.dll/xml and not PublisherApi.dll/xml)

### Bug Fixes
* Fix [#216](https://github.com/NetOfficeFw/NetOffice/issues/216) _Publisher NuGet Package has WordApi.dll/xml and not PublisherApi.dll/xml_

## v1.7.4.3

### Breaking Changes
* Changed method in `COMAddin` class: `virtual bool OnCreateTaskPaneInfo(TaskPaneInfo paneInfo)`
  > The meaning of the `bool` result has been changed and it is `true` by default now.
  > `true` means the Pane should have been created, otherwise `false`.

### Bug Fixes
* Fix [#193](https://github.com/NetOfficeFw/NetOffice/issues/193) _Can not get Addin from customArguments in OnConnection of ITaskpane_

## v1.7.4.2
* `COMAddin` supports custom addin object - see Word addin example **06 Custom Addin Object**

### Bug Fixes
* Fix [OSDN-37880](https://osdn.net/projects/netoffice/ticket/37880) _Underlying ribbon does not calling_
* Fix [OSDN-37747](https://osdn.net/projects/netoffice/ticket/37747) _DAO Fields_

## v1.7.4.1

### General
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

### Breaking Changes
* `COMObject` has been replaced by `ICOMObject` interface.
  You may have to change some event trigger code from `COMObject` to `ICOMObject`.
* Some native interop interfaces has been moved to `*.Native` namespace
