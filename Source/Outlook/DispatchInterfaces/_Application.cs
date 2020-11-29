using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _Application 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Application : COMObject
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
                    _type = typeof(_Application);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Application(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Application"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Class"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Session"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Parent"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Name"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Version"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.COMAddIns"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Explorers"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Explorers Explorers
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Explorers>(this, "Explorers");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Inspectors"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Inspectors Inspectors
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Inspectors>(this, "Inspectors");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.LanguageSettings"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(this, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.ProductCode"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ProductCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProductCode");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(this, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFeatureInstall>(this, "FeatureInstall");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FeatureInstall", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Reminders"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Reminders Reminders
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Reminders>(this, "Reminders");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.DefaultProfileName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string DefaultProfileName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultProfileName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.IsTrusted"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IsTrusted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsTrusted");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Assistance"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.TimeZones"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.TimeZones TimeZones
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.TimeZones>(this, "TimeZones", NetOffice.OutlookApi.TimeZones.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.PickerDialog"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OfficeApi.PickerDialog PickerDialog
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerDialog>(this, "PickerDialog", NetOffice.OfficeApi.PickerDialog.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.ActiveExplorer"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Explorer ActiveExplorer()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Explorer>(this, "ActiveExplorer");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.ActiveInspector"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Inspector ActiveInspector()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Inspector>(this, "ActiveInspector");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.CreateItem"/> </remarks>
		/// <param name="itemType">NetOffice.OutlookApi.Enums.OlItemType itemType</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object CreateItem(NetOffice.OutlookApi.Enums.OlItemType itemType)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateItem", itemType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.CreateItemFromTemplate"/> </remarks>
		/// <param name="templatePath">string templatePath</param>
		/// <param name="inFolder">optional object inFolder</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object CreateItemFromTemplate(string templatePath, object inFolder)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateItemFromTemplate", templatePath, inFolder);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.CreateItemFromTemplate"/> </remarks>
		/// <param name="templatePath">string templatePath</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object CreateItemFromTemplate(string templatePath)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateItemFromTemplate", templatePath);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.CreateObject"/> </remarks>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object CreateObject(string objectName)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateObject", objectName);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.GetNamespace"/> </remarks>
		/// <param name="type">string type</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace GetNamespace(string type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._NameSpace>(this, "GetNamespace", type);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.Quit(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.ActiveWindow"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object ActiveWindow()
		{
			return Factory.ExecuteVariantMethodGet(this, "ActiveWindow");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.CopyFile"/> </remarks>
		/// <param name="filePath">string filePath</param>
		/// <param name="destFolderPath">string destFolderPath</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public object CopyFile(string filePath, string destFolderPath)
		{
			return Factory.ExecuteVariantMethodGet(this, "CopyFile", filePath, destFolderPath);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.AdvancedSearch"/> </remarks>
		/// <param name="scope">string scope</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="searchSubFolders">optional object searchSubFolders</param>
		/// <param name="tag">optional object tag</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders, object tag)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(this, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope, filter, searchSubFolders, tag);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.AdvancedSearch"/> </remarks>
		/// <param name="scope">string scope</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(this, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.AdvancedSearch"/> </remarks>
		/// <param name="scope">string scope</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(this, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope, filter);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.AdvancedSearch"/> </remarks>
		/// <param name="scope">string scope</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="searchSubFolders">optional object searchSubFolders</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(this, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope, filter, searchSubFolders);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.IsSearchSynchronous"/> </remarks>
		/// <param name="lookInFolders">string lookInFolders</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public bool IsSearchSynchronous(string lookInFolders)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsSearchSynchronous", lookInFolders);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pvar">object pvar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void GetNewNickNames(object pvar)
		{
			 Factory.ExecuteMethod(this, "GetNewNickNames", pvar);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.GetObjectReference"/> </remarks>
		/// <param name="item">object item</param>
		/// <param name="referenceType">NetOffice.OutlookApi.Enums.OlReferenceType referenceType</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object GetObjectReference(object item, NetOffice.OutlookApi.Enums.OlReferenceType referenceType)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetObjectReference", item, referenceType);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Application.RefreshFormRegionDefinition"/> </remarks>
		/// <param name="regionName">string regionName</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public void RefreshFormRegionDefinition(string regionName)
		{
			 Factory.ExecuteMethod(this, "RefreshFormRegionDefinition", regionName);
		}

		#endregion

		#pragma warning restore
	}
}
