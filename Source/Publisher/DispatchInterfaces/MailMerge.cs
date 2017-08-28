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
 	public class MailMerge : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MailMerge(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailMerge(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMerge(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMergeDataSource DataSource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataSource>(this, "DataSource", NetOffice.PublisherApi.MailMergeDataSource.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Destination
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Destination");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool DocumentUpdating
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DocumentUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DocumentUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ShowSendToCustom
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShowSendToCustom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSendToCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool SuppressBlankLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SuppressBlankLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SuppressBlankLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool ViewMailMergeFieldCodes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewMailMergeFieldCodes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewMailMergeFieldCodes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 WizardState
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WizardState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WizardState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.EmailMergeEnvelope EmailMergeEnvelope
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.EmailMergeEnvelope>(this, "EmailMergeEnvelope", NetOffice.PublisherApi.EmailMergeEnvelope.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbMergeType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbMergeType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
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
		public void Execute10(bool pause)
		{
			 Factory.ExecuteMethod(this, "Execute10", pause);
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
		public void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", new object[]{ bstrDataSource, bstrConnect, bstrTable, fOpenExclusive, fNeverPrompt });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void OpenDataSource()
		{
			 Factory.ExecuteMethod(this, "OpenDataSource");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", bstrDataSource);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource, object bstrConnect)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", bstrDataSource, bstrConnect);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", bstrDataSource, bstrConnect, bstrTable);
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
		public void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive)
		{
			 Factory.ExecuteMethod(this, "OpenDataSource", bstrDataSource, bstrConnect, bstrTable, fOpenExclusive);
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
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizard()
		{
			 Factory.ExecuteMethod(this, "ShowWizard");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", showDocumentStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep, object showTemplateStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", showDocumentStep, showTemplateStep);
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
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", showDocumentStep, showTemplateStep, showDataStep);
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
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
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
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizard", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		/// <param name="filename">optional string Filename = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Execute(bool pause, object destination, object filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Execute", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, pause, destination, filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Execute(bool pause)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Execute", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, pause);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pause">bool pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Execute(bool pause, object destination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Document>(this, "Execute", NetOffice.PublisherApi.Document.LateBindingApiWrapperType, pause, destination);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		/// <param name="includedOnly">optional bool IncludedOnly = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportRecipientList(string filename, object fileType, object includedOnly)
		{
			 Factory.ExecuteMethod(this, "ExportRecipientList", filename, fileType, includedOnly);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportRecipientList(string filename)
		{
			 Factory.ExecuteMethod(this, "ExportRecipientList", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ExportRecipientList(string filename, object fileType)
		{
			 Factory.ExecuteMethod(this, "ExportRecipientList", filename, fileType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void CreateShortcut(string filename)
		{
			 Factory.ExecuteMethod(this, "CreateShortcut", filename);
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
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType, object iStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep, mergeType, iStep });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizardEx()
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", showDocumentStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", showDocumentStep, showTemplateStep);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", showDocumentStep, showTemplateStep, showDataStep);
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
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
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
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep });
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
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep });
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
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType)
		{
			 Factory.ExecuteMethod(this, "ShowWizardEx", new object[]{ showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep, mergeType });
		}

		#endregion

		#pragma warning restore
	}
}
