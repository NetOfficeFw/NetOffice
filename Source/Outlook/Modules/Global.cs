using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Outlook, 9,10,11,12,14,15,16
    ///</summary>
    [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(OutlookApi.Application))]
	public static class GlobalModule
	{
		#region Fields

		private static ICOMObject _instance;

        #endregion

        #region Internal Properties

        internal static ICOMObject Instance
        {
            get
            {
                return _instance;
            }
            set
            {
                if ((null == value) || (null == _instance))
                    _instance = value;
            }
        }

        internal static Core Factory
		{
			get
			{
				if(null != _instance)
					 return _instance.Factory;
			else
				return Core.Default;
			}
		}

		internal static Invoker Invoker
		{
			get
			{
				if(null != _instance)
					 return _instance.Invoker;
			else
				return Invoker.Default;
			}
		}

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868973.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._Application Application
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(_instance, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865581.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OutlookApi.Enums.OlObjectClass Class
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(_instance, "Class");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866436.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._NameSpace Session
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(_instance, "Session");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869381.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public static object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(_instance, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.Assistant Assistant
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(_instance, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868248.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860684.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static string Version
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870066.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.COMAddIns COMAddIns
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(_instance, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868795.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._Explorers Explorers
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Explorers>(_instance, "Explorers");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868935.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._Inspectors Inspectors
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Inspectors>(_instance, "Inspectors");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867217.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.LanguageSettings LanguageSettings
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(_instance, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869152.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static string ProductCode
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "ProductCode");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.AnswerWizard AnswerWizard
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(_instance, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFeatureInstall>(_instance, "FeatureInstall");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(_instance, "FeatureInstall", value);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862144.aspx </remarks>
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._Reminders Reminders
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Reminders>(_instance, "Reminders");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865059.aspx </remarks>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        public static string DefaultProfileName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(_instance, "DefaultProfileName");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864729.aspx </remarks>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        public static bool IsTrusted
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(_instance, "IsTrusted");
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861554.aspx </remarks>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        public static NetOffice.OfficeApi.IAssistance Assistance
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(_instance, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867696.aspx </remarks>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        public static NetOffice.OutlookApi.TimeZones TimeZones
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.TimeZones>(_instance, "TimeZones", NetOffice.OutlookApi.TimeZones.LateBindingApiWrapperType);
            }
        }

        /// <summary>
        /// SupportByVersion Outlook 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861549.aspx </remarks>
        [SupportByVersion("Outlook", 14, 15, 16)]
        public static NetOffice.OfficeApi.PickerDialog PickerDialog
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerDialog>(_instance, "PickerDialog", NetOffice.OfficeApi.PickerDialog.LateBindingApiWrapperType);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870017.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._Explorer ActiveExplorer()
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Explorer>(_instance, "ActiveExplorer");
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863939.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._Inspector ActiveInspector()
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Inspector>(_instance, "ActiveInspector");
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869635.aspx </remarks>
        /// <param name="itemType">NetOffice.OutlookApi.Enums.OlItemType itemType</param>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static object CreateItem(NetOffice.OutlookApi.Enums.OlItemType itemType)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CreateItem", itemType);
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865637.aspx </remarks>
        /// <param name="templatePath">string templatePath</param>
        /// <param name="inFolder">optional object inFolder</param>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static object CreateItemFromTemplate(string templatePath, object inFolder)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CreateItemFromTemplate", templatePath, inFolder);
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865637.aspx </remarks>
        /// <param name="templatePath">string templatePath</param>
        [CustomMethod]
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static object CreateItemFromTemplate(string templatePath)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CreateItemFromTemplate", templatePath);
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860724.aspx </remarks>
        /// <param name="objectName">string objectName</param>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static object CreateObject(string objectName)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CreateObject", objectName);
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865800.aspx </remarks>
        /// <param name="type">string type</param>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public static NetOffice.OutlookApi._NameSpace GetNamespace(string type)
        {
            return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._NameSpace>(_instance, "GetNamespace", type);
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866010.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static void Quit()
        {
            Factory.ExecuteMethod(_instance, "Quit");
        }

        /// <summary>
        /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865654.aspx </remarks>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        public static object ActiveWindow()
        {
            return Factory.ExecuteVariantMethodGet(_instance, "ActiveWindow");
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869462.aspx </remarks>
        /// <param name="filePath">string filePath</param>
        /// <param name="destFolderPath">string destFolderPath</param>
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static object CopyFile(string filePath, string destFolderPath)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "CopyFile", filePath, destFolderPath);
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
        /// <param name="scope">string scope</param>
        /// <param name="filter">optional object filter</param>
        /// <param name="searchSubFolders">optional object searchSubFolders</param>
        /// <param name="tag">optional object tag</param>
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders, object tag)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(_instance, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope, filter, searchSubFolders, tag);
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
        /// <param name="scope">string scope</param>
        [CustomMethod]
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OutlookApi.Search AdvancedSearch(string scope)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(_instance, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope);
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
        /// <param name="scope">string scope</param>
        /// <param name="filter">optional object filter</param>
        [CustomMethod]
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(_instance, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope, filter);
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx </remarks>
        /// <param name="scope">string scope</param>
        /// <param name="filter">optional object filter</param>
        /// <param name="searchSubFolders">optional object searchSubFolders</param>
        [CustomMethod]
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Search>(_instance, "AdvancedSearch", NetOffice.OutlookApi.Search.LateBindingApiWrapperType, scope, filter, searchSubFolders);
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869145.aspx </remarks>
        /// <param name="lookInFolders">string lookInFolders</param>
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static bool IsSearchSynchronous(string lookInFolders)
        {
            return Factory.ExecuteBoolMethodGet(_instance, "IsSearchSynchronous", lookInFolders);
        }

        /// <summary>
        /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pvar">object pvar</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        public static void GetNewNickNames(object pvar)
        {
            Factory.ExecuteMethod(_instance, "GetNewNickNames", pvar);
        }

        /// <summary>
        /// SupportByVersion Outlook 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864203.aspx </remarks>
        /// <param name="item">object item</param>
        /// <param name="referenceType">NetOffice.OutlookApi.Enums.OlReferenceType referenceType</param>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        public static object GetObjectReference(object item, NetOffice.OutlookApi.Enums.OlReferenceType referenceType)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "GetObjectReference", item, referenceType);
        }

        /// <summary>
        /// SupportByVersion Outlook 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863374.aspx </remarks>
        /// <param name="regionName">string regionName</param>
        [SupportByVersion("Outlook", 14, 15, 16)]
        public static void RefreshFormRegionDefinition(string regionName)
        {
            Factory.ExecuteMethod(_instance, "RefreshFormRegionDefinition", regionName);
        }

        #endregion
    }
}
