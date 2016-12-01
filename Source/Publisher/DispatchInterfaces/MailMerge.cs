using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface MailMerge 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class MailMerge : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.MailMergeDataSource DataSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataSource", paramsArray);
				NetOffice.PublisherApi.MailMergeDataSource newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.MailMergeDataSource.LateBindingApiWrapperType) as NetOffice.PublisherApi.MailMergeDataSource;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Destination
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Destination", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool DocumentUpdating
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentUpdating", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DocumentUpdating", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ShowSendToCustom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowSendToCustom", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowSendToCustom", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool SuppressBlankLines
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SuppressBlankLines", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SuppressBlankLines", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool ViewMailMergeFieldCodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ViewMailMergeFieldCodes", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ViewMailMergeFieldCodes", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 WizardState
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WizardState", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WizardState", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.EmailMergeEnvelope EmailMergeEnvelope
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmailMergeEnvelope", paramsArray);
				NetOffice.PublisherApi.EmailMergeEnvelope newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.EmailMergeEnvelope.LateBindingApiWrapperType) as NetOffice.PublisherApi.EmailMergeEnvelope;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbMergeType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbMergeType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Type", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pause">bool Pause</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void Execute10(bool pause)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pause);
			Invoker.Method(this, "Execute10", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		/// <param name="fNeverPrompt">optional Int32 fNeverPrompt = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrDataSource, bstrConnect, bstrTable, fOpenExclusive, fNeverPrompt);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void OpenDataSource()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrDataSource);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource, object bstrConnect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrDataSource, bstrConnect);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrDataSource, bstrConnect, bstrTable);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrDataSource">optional string bstrDataSource = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void OpenDataSource(object bstrDataSource, object bstrConnect, object bstrTable, object fOpenExclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrDataSource, bstrConnect, bstrTable, fOpenExclusive);
			Invoker.Method(this, "OpenDataSource", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep, object showTemplateStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizard(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep);
			Invoker.Method(this, "ShowWizard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pause">bool Pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		/// <param name="filename">optional string Filename = </param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Execute(bool pause, object destination, object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pause, destination, filename);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.PublisherApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Document.LateBindingApiWrapperType) as NetOffice.PublisherApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pause">bool Pause</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Execute(bool pause)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pause);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.PublisherApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Document.LateBindingApiWrapperType) as NetOffice.PublisherApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pause">bool Pause</param>
		/// <param name="destination">optional NetOffice.PublisherApi.Enums.PbMailMergeDestination Destination = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Document Execute(bool pause, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pause, destination);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.PublisherApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Document.LateBindingApiWrapperType) as NetOffice.PublisherApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		/// <param name="includedOnly">optional bool IncludedOnly = true</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportRecipientList(string filename, object fileType, object includedOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileType, includedOnly);
			Invoker.Method(this, "ExportRecipientList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportRecipientList(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "ExportRecipientList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="fileType">optional NetOffice.PublisherApi.Enums.PbRecipientListFileType FileType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ExportRecipientList(string filename, object fileType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, fileType);
			Invoker.Method(this, "ExportRecipientList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void CreateShortcut(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "CreateShortcut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		/// <param name="mergeType">optional NetOffice.PublisherApi.Enums.PbMergeType MergeType = 0</param>
		/// <param name="iStep">optional Int32 iStep = 1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType, object iStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep, mergeType, iStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showDocumentStep">optional bool ShowDocumentStep = true</param>
		/// <param name="showTemplateStep">optional bool ShowTemplateStep = false</param>
		/// <param name="showDataStep">optional bool ShowDataStep = true</param>
		/// <param name="showWriteStep">optional bool ShowWriteStep = true</param>
		/// <param name="showPreviewStep">optional bool ShowPreviewStep = true</param>
		/// <param name="showMergeStep">optional bool ShowMergeStep = true</param>
		/// <param name="mergeType">optional NetOffice.PublisherApi.Enums.PbMergeType MergeType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ShowWizardEx(object showDocumentStep, object showTemplateStep, object showDataStep, object showWriteStep, object showPreviewStep, object showMergeStep, object mergeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showDocumentStep, showTemplateStep, showDataStep, showWriteStep, showPreviewStep, showMergeStep, mergeType);
			Invoker.Method(this, "ShowWizardEx", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}