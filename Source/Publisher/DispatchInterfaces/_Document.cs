using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface _Document 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00021242-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.PublisherApi.Document))]
    public interface _Document : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string ActivePrinter { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Window ActiveWindow { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbColorMode ColorMode { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ColorScheme ColorScheme { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object DefaultTabStop { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool EnvelopeVisible { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string FullName { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.LayoutGuides LayoutGuides { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.MsoEnvelope MailEnvelope { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.MailMerge MailMerge { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.MasterPages MasterPages { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Pages Pages { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.PageSetup PageSetup { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string Path { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.PublisherApi.Enums.PbPersonalInfoSet PersonalInformationSet { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Plates Plates { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ReadOnly { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbDirectionType DocumentDirection { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool Saved { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbFileFormat SaveFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ScratchArea ScratchArea { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Selection Selection { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Stories Stories { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Tags Tags { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextStyles TextStyles { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool ViewBoundariesAndGuides { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ViewTwoPageSpread { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Wizard Wizard { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.View ActiveView { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.AdvancedPrintOptions AdvancedPrintOptions { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.BorderArts BorderArts { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool IsDataSourceConnected { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.FindReplace Find { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 UndoActionsAvailable { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 RedoActionsAvailable { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ViewHorizontalBaseLineGuides { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ViewVerticalBaseLineGuides { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbPublicationType PublicationType { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Sections Sections { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.WebNavigationBarSets WebNavigationBarSets { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool RemovePersonalInformation { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool PrintPageBackgrounds { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ColorsInUse ColorsInUse { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool IsWizard { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange SurplusShapes { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ViewBoundaries { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ViewGuides { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.BuildingBlocks AvailableBuildingBlocks { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Plates CreatePlateCollection(NetOffice.PublisherApi.Enums.PbColorMode mode);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void EnterColorMode10(NetOffice.PublisherApi.Enums.PbColorMode mode);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="tagName">string tagName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange FindShapesByTag(string tagName);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOut(object from, object to, object printToFile, object copies, object collate);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOut(object from);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOut(object from, object to);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOut(object from, object to, object printToFile);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOut(object from, object to, object printToFile, object copies);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		/// <param name="addToRecentFiles">optional bool AddToRecentFiles = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SaveAs(object filename, object format, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SaveAs();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SaveAs(object filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="format">optional NetOffice.PublisherApi.Enums.PbFileFormat Format = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SaveAs(object filename, object format);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="oh">Int32 oh</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void SelectID(Int32 oh);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void UndoClear();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void UpdateOLEObjects();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Undo(object count);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void Undo();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Redo(object count);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void Redo();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="actionName">string actionName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void BeginCustomUndoAction(string actionName);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void EndCustomUndoAction();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void WebPagePreview();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="value">NetOffice.PublisherApi.Enums.PbPublicationType value</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ConvertPublicationType(NetOffice.PublisherApi.Enums.PbPublicationType value);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		/// <param name="deleteExcessInks">optional bool DeleteExcessInks = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates, object deleteExcessInks);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="mode">NetOffice.PublisherApi.Enums.PbColorMode mode</param>
		/// <param name="plates">optional object plates</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void EnterColorMode(NetOffice.PublisherApi.Enums.PbColorMode mode, object plates);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx(object from, object to, object printToFile, object copies, object collate, object printStyle);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx(object from);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx(object from, object to);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx(object from, object to, object printToFile);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx(object from, object to, object printToFile, object copies);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void PrintOutEx(object from, object to, object printToFile, object copies, object collate);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard wizard</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard, object design);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizard wizard</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ChangeDocument(NetOffice.PublisherApi.Enums.PbWizard wizard);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetBusinessInformation(string name);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="externalExporter">optional object externalExporter</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.PublisherApi.Enums.PbFixedFormatType format</param>
		/// <param name="filename">string filename</param>
		/// <param name="intent">optional NetOffice.PublisherApi.Enums.PbFixedFormatIntent Intent = 3</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="colorDownsampleTarget">optional Int32 ColorDownsampleTarget = -1</param>
		/// <param name="colorDownsampleThreshold">optional Int32 ColorDownsampleThreshold = -1</param>
		/// <param name="oneBitDownsampleTarget">optional Int32 OneBitDownsampleTarget = -1</param>
		/// <param name="oneBitDownsampleThreshold">optional Int32 OneBitDownsampleThreshold = -1</param>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="copies">optional Int32 Copies = -1</param>
		/// <param name="collate">optional bool Collate = true</param>
		/// <param name="printStyle">optional NetOffice.PublisherApi.Enums.PbPrintStyle PrintStyle = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportAsFixedFormat(NetOffice.PublisherApi.Enums.PbFixedFormatType format, string filename, object intent, object includeDocumentProperties, object colorDownsampleTarget, object colorDownsampleThreshold, object oneBitDownsampleTarget, object oneBitDownsampleThreshold, object from, object to, object copies, object collate, object printStyle, object docStructureTags, object bitmapMissingFonts, object useISO19005_1);

		#endregion
	}
}
