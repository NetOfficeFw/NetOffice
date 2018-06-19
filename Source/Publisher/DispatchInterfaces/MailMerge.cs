using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface MailMerge 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("4FA84469-DD6A-42D4-979F-ED62ABBDF44D")]
	public interface MailMerge : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

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
		NetOffice.PublisherApi.MailMergeDataSource DataSource { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 Destination { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool DocumentUpdating { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string ShowSendToCustom { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool SuppressBlankLines { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		bool ViewMailMergeFieldCodes { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 WizardState { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.EmailMergeEnvelope EmailMergeEnvelope { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbMergeType Type { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void Execute10(bool pause);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		/// <param name="fNeverPrompt">optional Int32 fNeverPrompt = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenDataSource();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenDataSource(object bstrDataSource);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenDataSource(object bstrDataSource, object bstrConnect);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard(object showDocumentStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard(object showDocumentStep, object showTemplateStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		/// <param name="filename">optional string Filename = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Document Execute(bool pause, object destination, object filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Document Execute(bool pause);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Document Execute(bool pause, object destination);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		/// <param name="includedOnly">optional bool IncludedOnly = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportRecipientList(string filename, object fileType, object includedOnly);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportRecipientList(string filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ExportRecipientList(string filename, object fileType);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void CreateShortcut(string filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		/// <param name="mergeType">optional NetOffice.PublisherApi.Enums.PbMergeType MergeType = 0</param>
		/// <param name="iStep">optional Int32 iStep = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType, object iStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		/// <param name="mergeType">optional NetOffice.PublisherApi.Enums.PbMergeType MergeType = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType);

		#endregion
	}
}
