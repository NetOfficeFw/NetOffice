using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CoreServices;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Application
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00063001-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.Application))]
    public interface _Application : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868973.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865581.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866436.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869381.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Assistant Assistant { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868248.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860684.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string Version { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870066.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.COMAddIns COMAddIns { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868795.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Explorers Explorers { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868935.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Inspectors Inspectors { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867217.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.LanguageSettings LanguageSettings { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869152.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		string ProductCode { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.AnswerWizard AnswerWizard { get; }

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall { get; set; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862144.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Reminders Reminders { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865059.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		string DefaultProfileName { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864729.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		bool IsTrusted { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861554.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OfficeApi.IAssistance Assistance { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867696.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.TimeZones TimeZones { get; }

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861549.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		NetOffice.OfficeApi.PickerDialog PickerDialog { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870017.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Explorer ActiveExplorer();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863939.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Inspector ActiveInspector();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869635.aspx </remarks>
		/// <param name="itemType">NetOffice.OutlookApi.Enums.OlItemType itemType</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object CreateItem(NetOffice.OutlookApi.Enums.OlItemType itemType);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865637.aspx </remarks>
		/// <param name="templatePath">string templatePath</param>
		/// <param name="inFolder">optional object inFolder</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object CreateItemFromTemplate(string templatePath, object inFolder);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865637.aspx </remarks>
		/// <param name="templatePath">string templatePath</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object CreateItemFromTemplate(string templatePath);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860724.aspx </remarks>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object CreateObject(string objectName);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865800.aspx </remarks>
		/// <param name="type">string type</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace GetNamespace(string type);

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866010.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		void Quit();

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865654.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		object ActiveWindow();

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869462.aspx </remarks>
		/// <param name="filePath">string filePath</param>
		/// <param name="destFolderPath">string destFolderPath</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		object CopyFile(string filePath, string destFolderPath);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
		/// <param name="scope">string scope</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="searchSubFolders">optional object searchSubFolders</param>
		/// <param name="tag">optional object tag</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders, object tag);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
		/// <param name="scope">string scope</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Search AdvancedSearch(string scope);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
		/// <param name="scope">string scope</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
		/// <param name="scope">string scope</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="searchSubFolders">optional object searchSubFolders</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869145.aspx </remarks>
		/// <param name="lookInFolders">string lookInFolders</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		bool IsSearchSynchronous(string lookInFolders);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pvar">object pvar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		void GetNewNickNames(object pvar);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864203.aspx </remarks>
		/// <param name="item">object item</param>
		/// <param name="referenceType">NetOffice.OutlookApi.Enums.OlReferenceType referenceType</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		object GetObjectReference(object item, NetOffice.OutlookApi.Enums.OlReferenceType referenceType);

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863374.aspx </remarks>
		/// <param name="regionName">string regionName</param>
		[SupportByVersion("Outlook", 14,15,16)]
		void RefreshFormRegionDefinition(string regionName);

		#endregion
	}
}
