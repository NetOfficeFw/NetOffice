using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Access, 9,10,11,12,14,15,16
    ///</summary>
    [SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(AccessApi.Application))]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192087.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Application", paramsArray);
				NetOffice.AccessApi.Application newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.Application.LateBindingApiWrapperType) as NetOffice.AccessApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836400.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public static object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822407.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public static object CodeContextObject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CodeContextObject", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835352.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string MenuBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "MenuBar", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "MenuBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845319.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Int32 CurrentObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CurrentObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196795.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string CurrentObjectName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CurrentObjectName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837183.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Forms Forms
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Forms", paramsArray);
				NetOffice.AccessApi.Forms newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.Forms.LateBindingApiWrapperType) as NetOffice.AccessApi.Forms;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834339.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Reports Reports
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Reports", paramsArray);
				NetOffice.AccessApi.Reports newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.Reports.LateBindingApiWrapperType) as NetOffice.AccessApi.Reports;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835056.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Screen Screen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Screen", paramsArray);
				NetOffice.AccessApi.Screen newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.Screen.LateBindingApiWrapperType) as NetOffice.AccessApi.Screen;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845564.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DoCmd DoCmd
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "DoCmd", paramsArray);
				NetOffice.AccessApi.DoCmd newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.DoCmd.LateBindingApiWrapperType) as NetOffice.AccessApi.DoCmd;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195236.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string ShortcutMenuBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "ShortcutMenuBar", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "ShortcutMenuBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821493.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Visible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836033.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool UserControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "UserControl", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "UserControl", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821724.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.DAOApi.DBEngine DBEngine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "DBEngine", paramsArray);
				NetOffice.DAOApi.DBEngine newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.DAOApi.DBEngine.LateBindingApiWrapperType) as NetOffice.DAOApi.DBEngine;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821379.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Assistant", paramsArray);
				NetOffice.OfficeApi.Assistant newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType) as NetOffice.OfficeApi.Assistant;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835326.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.References References
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "References", paramsArray);
				NetOffice.AccessApi.References newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.References.LateBindingApiWrapperType) as NetOffice.AccessApi.References;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836265.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Modules Modules
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Modules", paramsArray);
				NetOffice.AccessApi.Modules newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.Modules.LateBindingApiWrapperType) as NetOffice.AccessApi.Modules;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "FileSearch", paramsArray);
				NetOffice.OfficeApi.FileSearch newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileSearch;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823044.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool IsCompiled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "IsCompiled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822476.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "VBE", paramsArray);
				NetOffice.VBIDEApi.VBE newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBE;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DataAccessPages DataAccessPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "DataAccessPages", paramsArray);
				NetOffice.AccessApi.DataAccessPages newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.DataAccessPages.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static string ADOConnectString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "ADOConnectString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193770.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CurrentProject CurrentProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CurrentProject", paramsArray);
				NetOffice.AccessApi.CurrentProject newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.CurrentProject.LateBindingApiWrapperType) as NetOffice.AccessApi.CurrentProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193230.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CurrentData CurrentData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CurrentData", paramsArray);
				NetOffice.AccessApi.CurrentData newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.CurrentData.LateBindingApiWrapperType) as NetOffice.AccessApi.CurrentData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197047.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CodeProject CodeProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CodeProject", paramsArray);
				NetOffice.AccessApi.CodeProject newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.CodeProject.LateBindingApiWrapperType) as NetOffice.AccessApi.CodeProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836912.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.CodeData CodeData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "CodeData", paramsArray);
				NetOffice.AccessApi.CodeData newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.CodeData.LateBindingApiWrapperType) as NetOffice.AccessApi.CodeData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.AccessApi.WizHook WizHook
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "WizHook", paramsArray);
				NetOffice.AccessApi.WizHook newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.WizHook.LateBindingApiWrapperType) as NetOffice.AccessApi.WizHook;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822077.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string ProductCode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "ProductCode", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822463.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "COMAddIns", paramsArray);
				NetOffice.OfficeApi.COMAddIns newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType) as NetOffice.OfficeApi.COMAddIns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194961.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DefaultWebOptions DefaultWebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "DefaultWebOptions", paramsArray);
				NetOffice.AccessApi.DefaultWebOptions newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.DefaultWebOptions.LateBindingApiWrapperType) as NetOffice.AccessApi.DefaultWebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836634.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "LanguageSettings", paramsArray);
				NetOffice.OfficeApi.LanguageSettings newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType) as NetOffice.OfficeApi.LanguageSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "AnswerWizard", paramsArray);
				NetOffice.OfficeApi.AnswerWizard newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType) as NetOffice.OfficeApi.AnswerWizard;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822721.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "FeatureInstall", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFeatureInstall)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "FeatureInstall", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static object VGXFrameInterval
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "VGXFrameInterval", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx
		/// </summary>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(_instance, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, dialogType);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx
		/// Alias for get_FileDialog
		/// </summary>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersion("Access", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		public static NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return get_FileDialog(dialogType);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845884.aspx
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static bool BrokenReference
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "BrokenReference", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195779.aspx
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Printers Printers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Printers", paramsArray);
				NetOffice.AccessApi.Printers newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.Printers.LateBindingApiWrapperType) as NetOffice.AccessApi.Printers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821394.aspx
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static NetOffice.AccessApi._Printer Printer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Printer", paramsArray);
				NetOffice.AccessApi._Printer newObject = Factory.CreateObjectFromComProxy(_instance,returnItem) as NetOffice.AccessApi._Printer;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "Printer", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "MsoDebugOptions", paramsArray);
				NetOffice.OfficeApi.MsoDebugOptions newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType) as NetOffice.OfficeApi.MsoDebugOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192859.aspx
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static string Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Version", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835096.aspx
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static Int32 Build
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Build", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191715.aspx
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.NewFile NewFileTaskPane
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "NewFileTaskPane", paramsArray);
				NetOffice.OfficeApi.NewFile newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType) as NetOffice.OfficeApi.NewFile;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845345.aspx
		/// </summary>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static NetOffice.AccessApi._AutoCorrect AutoCorrect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "AutoCorrect", paramsArray);
				NetOffice.AccessApi._AutoCorrect newObject = Factory.CreateObjectFromComProxy(_instance,returnItem) as NetOffice.AccessApi._AutoCorrect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193178.aspx
		/// </summary>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "AutomationSecurity", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoAutomationSecurity)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(_instance, "AutomationSecurity", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845034.aspx
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.MacroError MacroError
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "MacroError", paramsArray);
				NetOffice.AccessApi.MacroError newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.MacroError.LateBindingApiWrapperType) as NetOffice.AccessApi.MacroError;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192459.aspx
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.TempVars TempVars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "TempVars", paramsArray);
				NetOffice.AccessApi.TempVars newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.TempVars.LateBindingApiWrapperType) as NetOffice.AccessApi.TempVars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192450.aspx
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "Assistance", paramsArray);
				NetOffice.OfficeApi.IAssistance newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType) as NetOffice.OfficeApi.IAssistance;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837286.aspx
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		public static NetOffice.AccessApi.WebServices WebServices
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "WebServices", paramsArray);
				NetOffice.AccessApi.WebServices newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.WebServices.LateBindingApiWrapperType) as NetOffice.AccessApi.WebServices;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.AccessApi.LocalVars LocalVars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "LocalVars", paramsArray);
				NetOffice.AccessApi.LocalVars newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.LocalVars.LateBindingApiWrapperType) as NetOffice.AccessApi.LocalVars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj249062.aspx
		/// </summary>
		[SupportByVersion("Access", 14,15,16)]
		public static NetOffice.AccessApi.ReturnVars ReturnVars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(_instance, "ReturnVars", paramsArray);
				NetOffice.AccessApi.ReturnVars newObject = Factory.CreateKnownObjectFromComProxy(_instance,returnItem,NetOffice.AccessApi.ReturnVars.LateBindingApiWrapperType) as NetOffice.AccessApi.ReturnVars;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void NewCurrentDatabase(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object template</param>
		/// <param name="siteAddress">optional string SiteAddress = </param>
		/// <param name="listID">optional string ListID = </param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress, object listID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat, template, siteAddress, listID);
			Invoker.Method(_instance, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabase(string filepath, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat);
			Invoker.Method(_instance, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object template</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabase(string filepath, object fileFormat, object template)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat, template);
			Invoker.Method(_instance, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object template</param>
		/// <param name="siteAddress">optional string SiteAddress = </param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat, template, siteAddress);
			Invoker.Method(_instance, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void OpenCurrentDatabase(string filepath, object exclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive);
			Invoker.Method(_instance, "OpenCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		/// <param name="bstrPassword">optional string bstrPassword = </param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void OpenCurrentDatabase(string filepath, object exclusive, object bstrPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive, bstrPassword);
			Invoker.Method(_instance, "OpenCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void OpenCurrentDatabase(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "OpenCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192308.aspx
		/// </summary>
		/// <param name="optionName">string optionName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object GetOption(string optionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(optionName);
			object returnItem = Invoker.MethodReturn(_instance, "GetOption", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195513.aspx
		/// </summary>
		/// <param name="optionName">string optionName</param>
		/// <param name="setting">object setting</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void SetOption(string optionName, object setting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(optionName, setting);
			Invoker.Method(_instance, "SetOption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx
		/// </summary>
		/// <param name="echoOn">Int16 echoOn</param>
		/// <param name="bstrStatusBarText">optional string bstrStatusBarText = </param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Echo(Int16 echoOn, object bstrStatusBarText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(echoOn, bstrStatusBarText);
			Invoker.Method(_instance, "Echo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx
		/// </summary>
		/// <param name="echoOn">Int16 echoOn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Echo(Int16 echoOn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(echoOn);
			Invoker.Method(_instance, "Echo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836850.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void CloseCurrentDatabase()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "CloseCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx
		/// </summary>
		/// <param name="option">optional NetOffice.AccessApi.Enums.AcQuitOption Option = 1</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Quit(object option)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(option);
			Invoker.Method(_instance, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx
		/// </summary>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		/// <param name="argument2">optional object argument2</param>
		/// <param name="argument3">optional object argument3</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2, object argument3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, argument2, argument3);
			object returnItem = Invoker.MethodReturn(_instance, "SysCmd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx
		/// </summary>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action);
			object returnItem = Invoker.MethodReturn(_instance, "SysCmd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx
		/// </summary>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction action</param>
		/// <param name="argument2">optional object argument2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, argument2);
			object returnItem = Invoker.MethodReturn(_instance, "SysCmd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx
		/// </summary>
		/// <param name="database">optional object database</param>
		/// <param name="formTemplate">optional object formTemplate</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Form CreateForm(object database, object formTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database, formTemplate);
			object returnItem = Invoker.MethodReturn(_instance, "CreateForm", paramsArray);
			NetOffice.AccessApi.Form newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Form.LateBindingApiWrapperType) as NetOffice.AccessApi.Form;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Form CreateForm()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CreateForm", paramsArray);
			NetOffice.AccessApi.Form newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Form.LateBindingApiWrapperType) as NetOffice.AccessApi.Form;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx
		/// </summary>
		/// <param name="database">optional object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Form CreateForm(object database)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database);
			object returnItem = Invoker.MethodReturn(_instance, "CreateForm", paramsArray);
			NetOffice.AccessApi.Form newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Form.LateBindingApiWrapperType) as NetOffice.AccessApi.Form;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx
		/// </summary>
		/// <param name="database">optional object database</param>
		/// <param name="reportTemplate">optional object reportTemplate</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Report CreateReport(object database, object reportTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database, reportTemplate);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReport", paramsArray);
			NetOffice.AccessApi.Report newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Report.LateBindingApiWrapperType) as NetOffice.AccessApi.Report;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Report CreateReport()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CreateReport", paramsArray);
			NetOffice.AccessApi.Report newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Report.LateBindingApiWrapperType) as NetOffice.AccessApi.Report;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx
		/// </summary>
		/// <param name="database">optional object database</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Report CreateReport(object database)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReport", paramsArray);
			NetOffice.AccessApi.Report newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Report.LateBindingApiWrapperType) as NetOffice.AccessApi.Report;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlSource">string controlSource</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlEx(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, controlSource, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlEx", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlName">string controlName</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlEx(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, controlName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlEx", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836733.aspx
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlName">string controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DeleteControl(string formName, string controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlName);
			Invoker.Method(_instance, "DeleteControl", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191904.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlName">string controlName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DeleteReportControl(string reportName, string controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlName);
			Invoker.Method(_instance, "DeleteReportControl", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197044.aspx
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="expression">string expression</param>
		/// <param name="header">Int16 header</param>
		/// <param name="footer">Int16 footer</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Int32 CreateGroupLevel(string reportName, string expression, Int16 header, Int16 footer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, expression, header, footer);
			object returnItem = Invoker.MethodReturn(_instance, "CreateGroupLevel", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DMin(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DMin", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DMin(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DMin", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DMax(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DMax", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835050.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DMax(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DMax", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DSum(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DSum", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193998.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DSum(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DSum", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DAvg(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DAvg", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197744.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DAvg(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DAvg", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DLookup(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DLookup", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834404.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DLookup(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DLookup", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DLast(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DLast", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845086.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DLast(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DLast", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DVar(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DVar", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835667.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DVar(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DVar", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DVarP(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DVarP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197963.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DVarP(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DVarP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DStDev(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DStDev", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192869.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DStDev(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DStDev", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DStDevP(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DStDevP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834343.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DStDevP(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DStDevP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DFirst(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DFirst", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195230.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DFirst(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DFirst", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		/// <param name="criteria">optional object criteria</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DCount(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(_instance, "DCount", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191865.aspx
		/// </summary>
		/// <param name="expr">string expr</param>
		/// <param name="domain">string domain</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DCount(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(_instance, "DCount", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834705.aspx
		/// </summary>
		/// <param name="stringExpr">string stringExpr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Eval(string stringExpr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(stringExpr);
			object returnItem = Invoker.MethodReturn(_instance, "Eval", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845778.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string CurrentUser()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CurrentUser", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196189.aspx
		/// </summary>
		/// <param name="application">string application</param>
		/// <param name="topic">string topic</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object DDEInitiate(string application, string topic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(application, topic);
			object returnItem = Invoker.MethodReturn(_instance, "DDEInitiate", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197936.aspx
		/// </summary>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="command">string command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DDEExecute(object chanNum, string command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum, command);
			Invoker.Method(_instance, "DDEExecute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194752.aspx
		/// </summary>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="item">string item</param>
		/// <param name="data">string data</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DDEPoke(object chanNum, string item, string data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum, item, data);
			Invoker.Method(_instance, "DDEPoke", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823145.aspx
		/// </summary>
		/// <param name="chanNum">object chanNum</param>
		/// <param name="item">string item</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string DDERequest(object chanNum, string item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum, item);
			object returnItem = Invoker.MethodReturn(_instance, "DDERequest", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197795.aspx
		/// </summary>
		/// <param name="chanNum">object chanNum</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DDETerminate(object chanNum)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum);
			Invoker.Method(_instance, "DDETerminate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845193.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DDETerminateAll()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "DDETerminateAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835631.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.DAOApi.Database CurrentDb()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CurrentDb", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196457.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.DAOApi.Database CodeDb()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CodeDb", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="hwnd">Int32 hwnd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void BeginUndoable(Int32 hwnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hwnd);
			Invoker.Method(_instance, "BeginUndoable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="yesno">Int16 yesno</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void SetUndoRecording(Int16 yesno)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(yesno);
			Invoker.Method(_instance, "SetUndoRecording", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845070.aspx
		/// </summary>
		/// <param name="field">string field</param>
		/// <param name="fieldType">Int16 fieldType</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string BuildCriteria(string field, Int16 fieldType, string expression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, fieldType, expression);
			object returnItem = Invoker.MethodReturn(_instance, "BuildCriteria", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="moduleName">string moduleName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void InsertText(string text, string moduleName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, moduleName);
			Invoker.Method(_instance, "InsertText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void ReloadAddIns()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "ReloadAddIns", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836901.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.DAOApi.Workspace DefaultWorkspaceClone()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "DefaultWorkspaceClone", paramsArray);
			NetOffice.DAOApi.Workspace newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.DAOApi.Workspace.LateBindingApiWrapperType) as NetOffice.DAOApi.Workspace;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197957.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void RefreshTitleBar()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "RefreshTitleBar", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="changeFrom">string changeFrom</param>
		/// <param name="changeTo">string changeTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void AddAutoCorrect(string changeFrom, string changeTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(changeFrom, changeTo);
			Invoker.Method(_instance, "AddAutoCorrect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="changeFrom">string changeFrom</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void DelAutoCorrect(string changeFrom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(changeFrom);
			Invoker.Method(_instance, "DelAutoCorrect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196179.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Int32 hWndAccessApp()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "hWndAccessApp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		/// <param name="arg30">optional object arg30</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string procedure</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(_instance, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx
		/// </summary>
		/// <param name="value">object value</param>
		/// <param name="valueIfNull">optional object valueIfNull</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Nz(object value, object valueIfNull)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value, valueIfNull);
			object returnItem = Invoker.MethodReturn(_instance, "Nz", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195223.aspx
		/// </summary>
		/// <param name="value">object value</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object Nz(object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value);
			object returnItem = Invoker.MethodReturn(_instance, "Nz", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835072.aspx
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object LoadPicture(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(_instance, "LoadPicture", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(_instance,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objtyp">Int32 objtyp</param>
		/// <param name="moduleName">string moduleName</param>
		/// <param name="fileName">string fileName</param>
		/// <param name="token">Int32 token</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void ReplaceModule(Int32 objtyp, string moduleName, string fileName, Int32 token)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objtyp, moduleName, fileName, token);
			Invoker.Method(_instance, "ReplaceModule", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196488.aspx
		/// </summary>
		/// <param name="errorNumber">object errorNumber</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object AccessError(object errorNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(errorNumber);
			object returnItem = Invoker.MethodReturn(_instance, "AccessError", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object BuilderString()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "BuilderString", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193935.aspx
		/// </summary>
		/// <param name="guid">object guid</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object StringFromGUID(object guid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(guid);
			object returnItem = Invoker.MethodReturn(_instance, "StringFromGUID", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197675.aspx
		/// </summary>
		/// <param name="_string">object string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object GUIDFromString(object _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(_instance, "GUIDFromString", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="id">Int32 id</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static object AppLoadString(Int32 id)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id);
			object returnItem = Invoker.MethodReturn(_instance, "AppLoadString", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method);
			Invoker.Method(_instance, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void SaveAsText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(_instance, "SaveAsText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void LoadFromText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(_instance, "LoadFromText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823011.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void AddToFavorites()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "AddToFavorites", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194960.aspx
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void RefreshDatabaseWindow()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "RefreshDatabaseWindow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191909.aspx
		/// </summary>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand command</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void RunCommand(NetOffice.AccessApi.Enums.AcCommand command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(command);
			Invoker.Method(_instance, "RunCommand", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx
		/// </summary>
		/// <param name="hyperlink">object hyperlink</param>
		/// <param name="part">optional NetOffice.AccessApi.Enums.AcHyperlinkPart Part = 0</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string HyperlinkPart(object hyperlink, object part)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hyperlink, part);
			object returnItem = Invoker.MethodReturn(_instance, "HyperlinkPart", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx
		/// </summary>
		/// <param name="hyperlink">object hyperlink</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static string HyperlinkPart(object hyperlink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hyperlink);
			object returnItem = Invoker.MethodReturn(_instance, "HyperlinkPart", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821756.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static bool GetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			object returnItem = Invoker.MethodReturn(_instance, "GetHiddenAttribute", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822459.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fHidden">bool fHidden</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void SetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, bool fHidden)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fHidden);
			Invoker.Method(_instance, "SetHiddenAttribute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="createNewFile">optional bool CreateNewFile = true</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName, object createNewFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, createNewFile);
			object returnItem = Invoker.MethodReturn(_instance, "CreateDataAccessPage", paramsArray);
			NetOffice.AccessApi.DataAccessPage newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DataAccessPage CreateDataAccessPage()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CreateDataAccessPage", paramsArray);
			NetOffice.AccessApi.DataAccessPage newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(_instance, "CreateDataAccessPage", paramsArray);
			NetOffice.AccessApi.DataAccessPage newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void NewAccessProject(string filepath, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, connect);
			Invoker.Method(_instance, "NewAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void NewAccessProject(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "NewAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void OpenAccessProject(string filepath, object exclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive);
			Invoker.Method(_instance, "OpenAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void OpenAccessProject(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "OpenAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object connect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void CreateAccessProject(string filepath, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, connect);
			Invoker.Method(_instance, "CreateAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static void CreateAccessProject(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "CreateAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx
		/// </summary>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		/// <param name="fullPrecision">optional object fullPrecision</param>
		/// <param name="triangulationPrecision">optional object triangulationPrecision</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision);
			object returnItem = Invoker.MethodReturn(_instance, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx
		/// </summary>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency);
			object returnItem = Invoker.MethodReturn(_instance, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx
		/// </summary>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		/// <param name="fullPrecision">optional object fullPrecision</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public static Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency, fullPrecision);
			object returnItem = Invoker.MethodReturn(_instance, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void OpenCurrentDatabaseOld(string filepath, object exclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive);
			Invoker.Method(_instance, "OpenCurrentDatabaseOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void OpenCurrentDatabaseOld(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "OpenCurrentDatabaseOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		/// <param name="replace">optional bool Replace = false</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID, object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, company, workgroupID, replace);
			Invoker.Method(_instance, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void CreateNewWorkgroupFile()
		{
			object[] paramsArray = null;
			Invoker.Method(_instance, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void CreateNewWorkgroupFile(object path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(_instance, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void CreateNewWorkgroupFile(object path, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name);
			Invoker.Method(_instance, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void CreateNewWorkgroupFile(object path, object name, object company)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, company);
			Invoker.Method(_instance, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, company, workgroupID);
			Invoker.Method(_instance, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195103.aspx
		/// </summary>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void SetDefaultWorkgroupFile(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(_instance, "SetDefaultWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193465.aspx
		/// </summary>
		/// <param name="sourceFilename">string sourceFilename</param>
		/// <param name="destinationFilename">string destinationFilename</param>
		/// <param name="destinationFileFormat">NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ConvertAccessProject(string sourceFilename, string destinationFilename, NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceFilename, destinationFilename, destinationFileFormat);
			Invoker.Method(_instance, "ConvertAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx
		/// </summary>
		/// <param name="sourceFile">string sourceFile</param>
		/// <param name="destinationFile">string destinationFile</param>
		/// <param name="logFile">optional bool LogFile = false</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static bool CompactRepair(string sourceFile, string destinationFile, object logFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceFile, destinationFile, logFile);
			object returnItem = Invoker.MethodReturn(_instance, "CompactRepair", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx
		/// </summary>
		/// <param name="sourceFile">string sourceFile</param>
		/// <param name="destinationFile">string destinationFile</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static bool CompactRepair(string sourceFile, string destinationFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceFile, destinationFile);
			object returnItem = Invoker.MethodReturn(_instance, "CompactRepair", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional Int32 OtherFlags = 0</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional NetOffice.AccessApi.Enums.AcExportXMLOtherFlags OtherFlags = 0</param>
		/// <param name="whereCondition">optional string WhereCondition = </param>
		/// <param name="additionalData">optional object additionalData</param>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition, object additionalData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition, additionalData);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional NetOffice.AccessApi.Enums.AcExportXMLOtherFlags OtherFlags = 0</param>
		/// <param name="whereCondition">optional string WhereCondition = </param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition);
			Invoker.Method(_instance, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx
		/// </summary>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="importOptions">optional NetOffice.AccessApi.Enums.AcImportXMLOption ImportOptions = 1</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ImportXML(string dataSource, object importOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, importOptions);
			Invoker.Method(_instance, "ImportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx
		/// </summary>
		/// <param name="dataSource">string dataSource</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public static void ImportXML(string dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource);
			Invoker.Method(_instance, "ImportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional Int32 OtherFlags = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType</param>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding);
			Invoker.Method(_instance, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx
		/// </summary>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		/// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
		/// <param name="scriptOption">optional NetOffice.AccessApi.Enums.AcTransformXMLScriptOption ScriptOption = 1</param>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput, object scriptOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, transformSource, outputTarget, wellFormedXMLOutput, scriptOption);
			Invoker.Method(_instance, "TransformXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx
		/// </summary>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void TransformXML(string dataSource, string transformSource, string outputTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, transformSource, outputTarget);
			Invoker.Method(_instance, "TransformXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx
		/// </summary>
		/// <param name="dataSource">string dataSource</param>
		/// <param name="transformSource">string transformSource</param>
		/// <param name="outputTarget">string outputTarget</param>
		/// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, transformSource, outputTarget, wellFormedXMLOutput);
			Invoker.Method(_instance, "TransformXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834773.aspx
		/// </summary>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static NetOffice.AccessApi._AdditionalData CreateAdditionalData()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "CreateAdditionalData", paramsArray);
			NetOffice.AccessApi._AdditionalData newObject = Factory.CreateObjectFromComProxy(_instance,returnItem) as NetOffice.AccessApi._AdditionalData;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public static bool IsMemberSafe(Int32 dispid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dispid);
			object returnItem = Invoker.MethodReturn(_instance, "IsMemberSafe", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void NewCurrentDatabaseOld(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(_instance, "NewCurrentDatabaseOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object parent</param>
		/// <param name="columnName">optional object columnName</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="formName">string formName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlSource">string controlSource</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateControlExOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, controlSource, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateControlExOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="reportName">string reportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType controlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection section</param>
		/// <param name="parent">string parent</param>
		/// <param name="controlName">string controlName</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public static NetOffice.AccessApi.Control CreateReportControlExOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, controlName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(_instance, "CreateReportControlExOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(_instance, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx
		/// </summary>
		/// <param name="richText">object richText</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static string PlainText(object richText, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(richText, length);
			object returnItem = Invoker.MethodReturn(_instance, "PlainText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx
		/// </summary>
		/// <param name="richText">object richText</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static string PlainText(object richText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(richText);
			object returnItem = Invoker.MethodReturn(_instance, "PlainText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx
		/// </summary>
		/// <param name="plainText">object plainText</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static string HtmlEncode(object plainText, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(plainText, length);
			object returnItem = Invoker.MethodReturn(_instance, "HtmlEncode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx
		/// </summary>
		/// <param name="plainText">object plainText</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static string HtmlEncode(object plainText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(plainText);
			object returnItem = Invoker.MethodReturn(_instance, "HtmlEncode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194416.aspx
		/// </summary>
		/// <param name="customUIName">string customUIName</param>
		/// <param name="customUIXML">string customUIXML</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static void LoadCustomUI(string customUIName, string customUIXML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customUIName, customUIXML);
			Invoker.Method(_instance, "LoadCustomUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193467.aspx
		/// </summary>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ExportNavigationPane(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(_instance, "ExportNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="fAppendOnly">optional bool fAppendOnly = false</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ImportNavigationPane(string path, object fAppendOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fAppendOnly);
			Invoker.Method(_instance, "ImportNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx
		/// </summary>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ImportNavigationPane(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(_instance, "ImportNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835727.aspx
		/// </summary>
		/// <param name="tableName">string tableName</param>
		/// <param name="columnName">string columnName</param>
		/// <param name="queryString">string queryString</param>
		[SupportByVersion("Access", 12,14,15,16)]
		public static string ColumnHistory(string tableName, string columnName, string queryString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tableName, columnName, queryString);
			object returnItem = Invoker.MethodReturn(_instance, "ColumnHistory", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage, object toPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage, toPage);
			Invoker.Method(_instance, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType);
			Invoker.Method(_instance, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType, selectedRecords);
			Invoker.Method(_instance, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		/// <param name="externalExporter">object externalExporter</param>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType objectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 12,14,15,16)]
		public static void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage);
			Invoker.Method(_instance, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821429.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(_instance, "SaveAsAXL", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845765.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		public static void LoadFromAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(_instance, "LoadFromAXL", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		/// <param name="applicationPart">optional object applicationPart</param>
		/// <param name="includeData">optional object includeData</param>
		/// <param name="variation">optional object variation</param>
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData, object variation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData, variation);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		/// <param name="applicationPart">optional object applicationPart</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="title">string title</param>
		/// <param name="iconPath">string iconPath</param>
		/// <param name="coreTable">string coreTable</param>
		/// <param name="category">string category</param>
		/// <param name="previewPath">optional object previewPath</param>
		/// <param name="description">optional object description</param>
		/// <param name="instantiationForm">optional object instantiationForm</param>
		/// <param name="applicationPart">optional object applicationPart</param>
		/// <param name="includeData">optional object includeData</param>
		[CustomMethod]
		[SupportByVersion("Access", 14,15,16)]
		public static void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData);
			Invoker.Method(_instance, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835421.aspx
		/// </summary>
		/// <param name="path">string path</param>
		[SupportByVersion("Access", 14,15,16)]
		public static void InstantiateTemplate(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(_instance, "InstantiateTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834388.aspx
		/// </summary>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption</param>
		[SupportByVersion("Access", 14,15,16)]
		public static object CurrentWebUser(NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(displayOption);
			object returnItem = Invoker.MethodReturn(_instance, "CurrentWebUser", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836539.aspx
		/// </summary>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption</param>
		[SupportByVersion("Access", 14,15,16)]
		public static object CurrentWebUserGroups(NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(displayOption);
			object returnItem = Invoker.MethodReturn(_instance, "CurrentWebUserGroups", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(_instance, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193453.aspx
		/// </summary>
		/// <param name="groupNameOrID">object groupNameOrID</param>
		[SupportByVersion("Access", 14,15,16)]
		public static bool IsCurrentWebUserInGroup(object groupNameOrID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupNameOrID);
			object returnItem = Invoker.MethodReturn(_instance, "IsCurrentWebUserInGroup", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834368.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 14,15,16)]
		public static void DirtyObject(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(_instance, "DirtyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 14,15,16)]
		public static bool IsClient()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(_instance, "IsClient", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
	}
}
