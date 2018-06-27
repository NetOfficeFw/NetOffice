using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface MailMerge 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class MailMerge : COMObject, NetOffice.PublisherApi.MailMerge
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PublisherApi.MailMerge);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(MailMerge);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MailMerge() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeDataSource DataSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataSource>(this, "DataSource", typeof(NetOffice.PublisherApi.MailMergeDataSource));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 Destination
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Destination");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool DocumentUpdating
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DocumentUpdating");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DocumentUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ShowSendToCustom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShowSendToCustom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowSendToCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool SuppressBlankLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SuppressBlankLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SuppressBlankLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool ViewMailMergeFieldCodes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewMailMergeFieldCodes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewMailMergeFieldCodes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 WizardState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WizardState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WizardState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.EmailMergeEnvelope EmailMergeEnvelope
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.EmailMergeEnvelope>(this, "EmailMergeEnvelope", typeof(NetOffice.PublisherApi.EmailMergeEnvelope));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbMergeType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbMergeType>(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Execute10(bool pause)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Execute10", pause);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		/// <param name="fNeverPrompt">optional Int32 fNeverPrompt = 1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", new object[]{ bstrDataSource, bstrConnect, bstrTable, fOpenExclusive, fNeverPrompt });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenDataSource()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenDataSource(object bstrDataSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", bstrDataSource);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenDataSource(object bstrDataSource, object bstrConnect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", bstrDataSource, bstrConnect);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", bstrDataSource, bstrConnect, bstrTable);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenDataSource", bstrDataSource, bstrConnect, bstrTable, fOpenExclusive);
		}

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
		public virtual void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizard()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizard(object showDocumentStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", showDocumentStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizard(object showDocumentStep, object showTemplateStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", showDocumentStep, showTemplateStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", showDocumentStep, showTemplateStep, showDataStep);
		}

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
		public virtual void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
		}

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
		public virtual void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizard", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		/// <param name="filename">optional string Filename = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Execute(bool pause, object destination, object filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Execute", typeof(NetOffice.PublisherApi.Document), pause, destination, filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Execute(bool pause)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Execute", typeof(NetOffice.PublisherApi.Document), pause);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Document Execute(bool pause, object destination)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Execute", typeof(NetOffice.PublisherApi.Document), pause, destination);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		/// <param name="includedOnly">optional bool IncludedOnly = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ExportRecipientList(string filename, object fileType, object includedOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportRecipientList", filename, fileType, includedOnly);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ExportRecipientList(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportRecipientList", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ExportRecipientList(string filename, object fileType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportRecipientList", filename, fileType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void CreateShortcut(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateShortcut", filename);
		}

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
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType, object iStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep, mergeType, iStep });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardEx()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardEx(object showDocumentStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", showDocumentStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", showDocumentStep, showTemplateStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", showDocumentStep, showTemplateStep, showDataStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
		}

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
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep });
		}

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
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep });
		}

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
		public virtual void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep, mergeType });
		}

		#endregion

		#pragma warning restore
	}
}


