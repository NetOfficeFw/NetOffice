using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.CoreServices;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.PowerPointApi
{
    /// <summary>
    /// DispatchInterface _Application
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("91493442-5A91-11CF-8700-00AA0060263B")]
    [CoClassSource(typeof(NetOffice.PowerPointApi.Application))]
	public interface _Application : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746387.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentations Presentations { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746218.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.DocumentWindows Windows { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.PowerPointApi.PPDialogs Dialogs { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745295.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.DocumentWindow ActiveWindow { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744912.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Presentation ActivePresentation { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744816.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.SlideShowWindows SlideShowWindows { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744604.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.CommandBars CommandBars { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745905.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Path { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746231.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746732.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Assistant Assistant { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.FileSearch FileSearch { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.IFind FileFind { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746539.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Build { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746225.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744908.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string OperatingSystem { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744770.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string ActivePrinter { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744619.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744952.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.AddIns AddIns { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744521.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.VBIDEApi.VBE VBE { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745458.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Left { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746410.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Top { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746076.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Width { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744702.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Single Height { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744049.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpWindowState WindowState { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745566.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
		[KnownIssue]
		[Obsolete("Int32 HWND is unavailable. Use ApplicationUtils instead.")]
        Int32 HWND { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745671.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Active { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.AnswerWizard AnswerWizard { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746702.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.COMAddIns COMAddIns { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744276.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string ProductCode { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.DefaultWebOptions DefaultWebOptions { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745687.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.LanguageSettings LanguageSettings { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745929.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState ShowWindowsInTaskbar { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.PowerPointApi.Marker Marker { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744258.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744016.aspx </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744016.aspx </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746016.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState DisplayGridLines { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745661.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745695.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.NewFile NewPresentation { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746503.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpAlertLevel DisplayAlerts { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745925.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState ShowStartupDialog { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744774.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.AutoCorrect AutoCorrect { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744854.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Options Options { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744758.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		bool DisplayDocumentInformationPanel { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743833.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.IAssistance Assistance { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745260.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		Int32 ActiveEncryptionSession { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744365.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.FileConverters FileConverters { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743963.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745345.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745159.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.SmartArtColors SmartArtColors { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744225.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.ProtectedViewWindows ProtectedViewWindows { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746155.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.ProtectedViewWindow ActiveProtectedViewWindow { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746169.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool IsSandboxed { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.PowerPointApi.ResampleMediaTasks ResampleMediaTasks { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745623.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229713.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		bool ChartDataPointTrack { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228516.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.OfficeApi.Enums.MsoTriState DisplayGuides { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745709.aspx </remarks>
		/// <param name="helpFile">optional string HelpFile = vbappt9.chm</param>
		/// <param name="contextID">optional Int32 ContextID = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Help(object helpFile, object contextID);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745709.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Help();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745709.aspx </remarks>
		/// <param name="helpFile">optional string HelpFile = vbappt9.chm</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Help(object helpFile);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746388.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744221.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="safeArrayOfParams">optional object[] safeArrayOfParams</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		object Run(string macroName, object[] safeArrayOfParams);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744221.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		object Run(string macroName);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpFileDialogType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.FileDialog FileDialog(NetOffice.PowerPointApi.Enums.PpFileDialogType type);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pWindow">NetOffice.PowerPointApi.DocumentWindow pWindow</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void LaunchSpelling(NetOffice.PowerPointApi.DocumentWindow pWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745072.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Activate();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="persist">optional bool Persist = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		bool GetOptionFlag(Int32 option, object persist);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		bool GetOptionFlag(Int32 option);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="state">bool state</param>
		/// <param name="persist">optional bool Persist = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SetOptionFlag(Int32 option, bool state, object persist);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="state">bool state</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SetOptionFlag(Int32 option, bool state);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpFileDialogType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		object PPFileDialog(NetOffice.PowerPointApi.Enums.PpFileDialogType type);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="marker">Int32 marker</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SetPerfMarker(Int32 marker);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void LaunchPublishSlidesDialog(string slideLibraryUrl);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="slideUrls">object slideUrls</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void LaunchSendToPPTDialog(object slideUrls);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745395.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void StartNewUndoEntry();

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229517.aspx </remarks>
		/// <param name="themeFileName">string themeFileName</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Theme OpenThemeFile(string themeFileName);

		#endregion
	}
}
