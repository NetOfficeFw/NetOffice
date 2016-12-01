using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _Application 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Application : COMObject
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
                    _type = typeof(_Application);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192087.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.AccessApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Application.LateBindingApiWrapperType) as NetOffice.AccessApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836400.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822407.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object CodeContextObject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodeContextObject", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835352.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string MenuBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MenuBar", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MenuBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845319.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 CurrentObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196795.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string CurrentObjectName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentObjectName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837183.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Forms Forms
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Forms", paramsArray);
				NetOffice.AccessApi.Forms newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Forms.LateBindingApiWrapperType) as NetOffice.AccessApi.Forms;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834339.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Reports Reports
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Reports", paramsArray);
				NetOffice.AccessApi.Reports newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Reports.LateBindingApiWrapperType) as NetOffice.AccessApi.Reports;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835056.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Screen Screen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Screen", paramsArray);
				NetOffice.AccessApi.Screen newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Screen.LateBindingApiWrapperType) as NetOffice.AccessApi.Screen;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845564.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.DoCmd DoCmd
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DoCmd", paramsArray);
				NetOffice.AccessApi.DoCmd newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.DoCmd.LateBindingApiWrapperType) as NetOffice.AccessApi.DoCmd;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195236.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string ShortcutMenuBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShortcutMenuBar", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShortcutMenuBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821493.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836033.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool UserControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserControl", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserControl", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821724.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.DAOApi.DBEngine DBEngine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DBEngine", paramsArray);
				NetOffice.DAOApi.DBEngine newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.DBEngine.LateBindingApiWrapperType) as NetOffice.DAOApi.DBEngine;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821379.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assistant", paramsArray);
				NetOffice.OfficeApi.Assistant newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType) as NetOffice.OfficeApi.Assistant;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835326.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.References References
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "References", paramsArray);
				NetOffice.AccessApi.References newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.References.LateBindingApiWrapperType) as NetOffice.AccessApi.References;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836265.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Modules Modules
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Modules", paramsArray);
				NetOffice.AccessApi.Modules newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Modules.LateBindingApiWrapperType) as NetOffice.AccessApi.Modules;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileSearch", paramsArray);
				NetOffice.OfficeApi.FileSearch newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileSearch;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823044.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool IsCompiled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsCompiled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822476.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBE", paramsArray);
				NetOffice.VBIDEApi.VBE newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBE;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.DataAccessPages DataAccessPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataAccessPages", paramsArray);
				NetOffice.AccessApi.DataAccessPages newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.DataAccessPages.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ADOConnectString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ADOConnectString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193770.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.CurrentProject CurrentProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentProject", paramsArray);
				NetOffice.AccessApi.CurrentProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.CurrentProject.LateBindingApiWrapperType) as NetOffice.AccessApi.CurrentProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193230.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.CurrentData CurrentData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentData", paramsArray);
				NetOffice.AccessApi.CurrentData newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.CurrentData.LateBindingApiWrapperType) as NetOffice.AccessApi.CurrentData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197047.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.CodeProject CodeProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodeProject", paramsArray);
				NetOffice.AccessApi.CodeProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.CodeProject.LateBindingApiWrapperType) as NetOffice.AccessApi.CodeProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836912.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.CodeData CodeData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodeData", paramsArray);
				NetOffice.AccessApi.CodeData newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.CodeData.LateBindingApiWrapperType) as NetOffice.AccessApi.CodeData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.WizHook WizHook
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WizHook", paramsArray);
				NetOffice.AccessApi.WizHook newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.WizHook.LateBindingApiWrapperType) as NetOffice.AccessApi.WizHook;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822077.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string ProductCode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProductCode", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822463.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "COMAddIns", paramsArray);
				NetOffice.OfficeApi.COMAddIns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType) as NetOffice.OfficeApi.COMAddIns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194961.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.DefaultWebOptions DefaultWebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultWebOptions", paramsArray);
				NetOffice.AccessApi.DefaultWebOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.DefaultWebOptions.LateBindingApiWrapperType) as NetOffice.AccessApi.DefaultWebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836634.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageSettings", paramsArray);
				NetOffice.OfficeApi.LanguageSettings newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType) as NetOffice.OfficeApi.LanguageSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AnswerWizard", paramsArray);
				NetOffice.OfficeApi.AnswerWizard newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType) as NetOffice.OfficeApi.AnswerWizard;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822721.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FeatureInstall", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFeatureInstall)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FeatureInstall", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object VGXFrameInterval
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VGXFrameInterval", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(dialogType);
			object returnItem = Invoker.PropertyGet(this, "FileDialog", paramsArray);
			NetOffice.OfficeApi.FileDialog newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196794.aspx
		/// Alias for get_FileDialog
		/// </summary>
		/// <param name="dialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType dialogType)
		{
			return get_FileDialog(dialogType);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845884.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool BrokenReference
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BrokenReference", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195779.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Printers Printers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Printers", paramsArray);
				NetOffice.AccessApi.Printers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Printers.LateBindingApiWrapperType) as NetOffice.AccessApi.Printers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821394.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi._Printer Printer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Printer", paramsArray);
				NetOffice.AccessApi._Printer newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._Printer;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Printer", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MsoDebugOptions", paramsArray);
				NetOffice.OfficeApi.MsoDebugOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType) as NetOffice.OfficeApi.MsoDebugOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192859.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public string Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Version", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835096.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public Int32 Build
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Build", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191715.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.NewFile NewFileTaskPane
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NewFileTaskPane", paramsArray);
				NetOffice.OfficeApi.NewFile newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType) as NetOffice.OfficeApi.NewFile;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845345.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public NetOffice.AccessApi._AutoCorrect AutoCorrect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoCorrect", paramsArray);
				NetOffice.AccessApi._AutoCorrect newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._AutoCorrect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193178.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutomationSecurity", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoAutomationSecurity)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutomationSecurity", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845034.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.MacroError MacroError
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MacroError", paramsArray);
				NetOffice.AccessApi.MacroError newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.MacroError.LateBindingApiWrapperType) as NetOffice.AccessApi.MacroError;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192459.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.TempVars TempVars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TempVars", paramsArray);
				NetOffice.AccessApi.TempVars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.TempVars.LateBindingApiWrapperType) as NetOffice.AccessApi.TempVars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192450.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assistance", paramsArray);
				NetOffice.OfficeApi.IAssistance newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType) as NetOffice.OfficeApi.IAssistance;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837286.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public NetOffice.AccessApi.WebServices WebServices
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebServices", paramsArray);
				NetOffice.AccessApi.WebServices newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.WebServices.LateBindingApiWrapperType) as NetOffice.AccessApi.WebServices;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.LocalVars LocalVars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LocalVars", paramsArray);
				NetOffice.AccessApi.LocalVars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.LocalVars.LateBindingApiWrapperType) as NetOffice.AccessApi.LocalVars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj249062.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public NetOffice.AccessApi.ReturnVars ReturnVars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReturnVars", paramsArray);
				NetOffice.AccessApi.ReturnVars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.ReturnVars.LateBindingApiWrapperType) as NetOffice.AccessApi.ReturnVars;
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
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void NewCurrentDatabase(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object Template</param>
		/// <param name="siteAddress">optional string SiteAddress = </param>
		/// <param name="listID">optional string ListID = </param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress, object listID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat, template, siteAddress, listID);
			Invoker.Method(this, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NewCurrentDatabase(string filepath, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat);
			Invoker.Method(this, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object Template</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NewCurrentDatabase(string filepath, object fileFormat, object template)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat, template);
			Invoker.Method(this, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195271.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="fileFormat">optional NetOffice.AccessApi.Enums.AcNewDatabaseFormat FileFormat = 0</param>
		/// <param name="template">optional object Template</param>
		/// <param name="siteAddress">optional string SiteAddress = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NewCurrentDatabase(string filepath, object fileFormat, object template, object siteAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, fileFormat, template, siteAddress);
			Invoker.Method(this, "NewCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenCurrentDatabase(string filepath, object exclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive);
			Invoker.Method(this, "OpenCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		/// <param name="bstrPassword">optional string bstrPassword = </param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenCurrentDatabase(string filepath, object exclusive, object bstrPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive, bstrPassword);
			Invoker.Method(this, "OpenCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837226.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenCurrentDatabase(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "OpenCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192308.aspx
		/// </summary>
		/// <param name="optionName">string OptionName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object GetOption(string optionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(optionName);
			object returnItem = Invoker.MethodReturn(this, "GetOption", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="optionName">string OptionName</param>
		/// <param name="setting">object Setting</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetOption(string optionName, object setting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(optionName, setting);
			Invoker.Method(this, "SetOption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx
		/// </summary>
		/// <param name="echoOn">Int16 EchoOn</param>
		/// <param name="bstrStatusBarText">optional string bstrStatusBarText = </param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Echo(Int16 echoOn, object bstrStatusBarText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(echoOn, bstrStatusBarText);
			Invoker.Method(this, "Echo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834500.aspx
		/// </summary>
		/// <param name="echoOn">Int16 EchoOn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Echo(Int16 echoOn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(echoOn);
			Invoker.Method(this, "Echo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836850.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CloseCurrentDatabase()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CloseCurrentDatabase", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx
		/// </summary>
		/// <param name="option">optional NetOffice.AccessApi.Enums.AcQuitOption Option = 1</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Quit(object option)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(option);
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844963.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193809.aspx
		/// </summary>
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction Action</param>
		/// <param name="argument2">optional object Argument2</param>
		/// <param name="argument3">optional object Argument3</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2, object argument3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, argument2, argument3);
			object returnItem = Invoker.MethodReturn(this, "SysCmd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction Action</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action);
			object returnItem = Invoker.MethodReturn(this, "SysCmd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="action">NetOffice.AccessApi.Enums.AcSysCmdAction Action</param>
		/// <param name="argument2">optional object Argument2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object SysCmd(NetOffice.AccessApi.Enums.AcSysCmdAction action, object argument2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, argument2);
			object returnItem = Invoker.MethodReturn(this, "SysCmd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="database">optional object Database</param>
		/// <param name="formTemplate">optional object FormTemplate</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Form CreateForm(object database, object formTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database, formTemplate);
			object returnItem = Invoker.MethodReturn(this, "CreateForm", paramsArray);
			NetOffice.AccessApi.Form newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Form.LateBindingApiWrapperType) as NetOffice.AccessApi.Form;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Form CreateForm()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateForm", paramsArray);
			NetOffice.AccessApi.Form newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Form.LateBindingApiWrapperType) as NetOffice.AccessApi.Form;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845361.aspx
		/// </summary>
		/// <param name="database">optional object Database</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Form CreateForm(object database)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database);
			object returnItem = Invoker.MethodReturn(this, "CreateForm", paramsArray);
			NetOffice.AccessApi.Form newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Form.LateBindingApiWrapperType) as NetOffice.AccessApi.Form;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx
		/// </summary>
		/// <param name="database">optional object Database</param>
		/// <param name="reportTemplate">optional object ReportTemplate</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Report CreateReport(object database, object reportTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database, reportTemplate);
			object returnItem = Invoker.MethodReturn(this, "CreateReport", paramsArray);
			NetOffice.AccessApi.Report newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Report.LateBindingApiWrapperType) as NetOffice.AccessApi.Report;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Report CreateReport()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateReport", paramsArray);
			NetOffice.AccessApi.Report newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Report.LateBindingApiWrapperType) as NetOffice.AccessApi.Report;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193499.aspx
		/// </summary>
		/// <param name="database">optional object Database</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Report CreateReport(object database)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(database);
			object returnItem = Invoker.MethodReturn(this, "CreateReport", paramsArray);
			NetOffice.AccessApi.Report newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Report.LateBindingApiWrapperType) as NetOffice.AccessApi.Report;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836740.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControl(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "CreateControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193518.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControl(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControl", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection Section</param>
		/// <param name="parent">string Parent</param>
		/// <param name="controlSource">string ControlSource</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlEx(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, controlSource, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateControlEx", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection Section</param>
		/// <param name="parent">string Parent</param>
		/// <param name="controlName">string ControlName</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlEx(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, controlName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlEx", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836733.aspx
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlName">string ControlName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DeleteControl(string formName, string controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlName);
			Invoker.Method(this, "DeleteControl", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191904.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlName">string ControlName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DeleteReportControl(string reportName, string controlName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlName);
			Invoker.Method(this, "DeleteReportControl", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197044.aspx
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="expression">string Expression</param>
		/// <param name="header">Int16 Header</param>
		/// <param name="footer">Int16 Footer</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 CreateGroupLevel(string reportName, string expression, Int16 header, Int16 footer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, expression, header, footer);
			object returnItem = Invoker.MethodReturn(this, "CreateGroupLevel", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834804.aspx
		/// </summary>
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DMin(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DMin", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DMin(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DMin", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DMax(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DMax", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DMax(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DMax", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DSum(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DSum", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DSum(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DSum", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DAvg(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DAvg", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DAvg(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DAvg", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DLookup(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DLookup", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DLookup(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DLookup", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DLast(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DLast", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DLast(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DLast", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DVar(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DVar", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DVar(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DVar", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DVarP(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DVarP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DVarP(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DVarP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DStDev(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DStDev", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DStDev(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DStDev", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DStDevP(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DStDevP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DStDevP(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DStDevP", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DFirst(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DFirst", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DFirst(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DFirst", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		/// <param name="criteria">optional object Criteria</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DCount(string expr, string domain, object criteria)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain, criteria);
			object returnItem = Invoker.MethodReturn(this, "DCount", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="expr">string Expr</param>
		/// <param name="domain">string Domain</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DCount(string expr, string domain)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expr, domain);
			object returnItem = Invoker.MethodReturn(this, "DCount", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="stringExpr">string StringExpr</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Eval(string stringExpr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(stringExpr);
			object returnItem = Invoker.MethodReturn(this, "Eval", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string CurrentUser()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CurrentUser", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196189.aspx
		/// </summary>
		/// <param name="application">string Application</param>
		/// <param name="topic">string Topic</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object DDEInitiate(string application, string topic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(application, topic);
			object returnItem = Invoker.MethodReturn(this, "DDEInitiate", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="chanNum">object ChanNum</param>
		/// <param name="command">string Command</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DDEExecute(object chanNum, string command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum, command);
			Invoker.Method(this, "DDEExecute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194752.aspx
		/// </summary>
		/// <param name="chanNum">object ChanNum</param>
		/// <param name="item">string Item</param>
		/// <param name="data">string Data</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DDEPoke(object chanNum, string item, string data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum, item, data);
			Invoker.Method(this, "DDEPoke", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823145.aspx
		/// </summary>
		/// <param name="chanNum">object ChanNum</param>
		/// <param name="item">string Item</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string DDERequest(object chanNum, string item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum, item);
			object returnItem = Invoker.MethodReturn(this, "DDERequest", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197795.aspx
		/// </summary>
		/// <param name="chanNum">object ChanNum</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DDETerminate(object chanNum)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(chanNum);
			Invoker.Method(this, "DDETerminate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845193.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DDETerminateAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DDETerminateAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835631.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.DAOApi.Database CurrentDb()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CurrentDb", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196457.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.DAOApi.Database CodeDb()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CodeDb", paramsArray);
			NetOffice.DAOApi.Database newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Database.LateBindingApiWrapperType) as NetOffice.DAOApi.Database;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="hwnd">Int32 Hwnd</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void BeginUndoable(Int32 hwnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hwnd);
			Invoker.Method(this, "BeginUndoable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="yesno">Int16 yesno</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetUndoRecording(Int16 yesno)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(yesno);
			Invoker.Method(this, "SetUndoRecording", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845070.aspx
		/// </summary>
		/// <param name="field">string Field</param>
		/// <param name="fieldType">Int16 FieldType</param>
		/// <param name="expression">string Expression</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string BuildCriteria(string field, Int16 fieldType, string expression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, fieldType, expression);
			object returnItem = Invoker.MethodReturn(this, "BuildCriteria", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="moduleName">string ModuleName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void InsertText(string text, string moduleName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, moduleName);
			Invoker.Method(this, "InsertText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ReloadAddIns()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReloadAddIns", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836901.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.DAOApi.Workspace DefaultWorkspaceClone()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DefaultWorkspaceClone", paramsArray);
			NetOffice.DAOApi.Workspace newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.Workspace.LateBindingApiWrapperType) as NetOffice.DAOApi.Workspace;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197957.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RefreshTitleBar()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshTitleBar", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="changeFrom">string ChangeFrom</param>
		/// <param name="changeTo">string ChangeTo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void AddAutoCorrect(string changeFrom, string changeTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(changeFrom, changeTo);
			Invoker.Method(this, "AddAutoCorrect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="changeFrom">string ChangeFrom</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void DelAutoCorrect(string changeFrom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(changeFrom);
			Invoker.Method(this, "DelAutoCorrect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196179.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Int32 hWndAccessApp()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "hWndAccessApp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193559.aspx
		/// </summary>
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="procedure">string Procedure</param>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Run(string procedure, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(procedure, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="value">object Value</param>
		/// <param name="valueIfNull">optional object ValueIfNull</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Nz(object value, object valueIfNull)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value, valueIfNull);
			object returnItem = Invoker.MethodReturn(this, "Nz", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="value">object Value</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object Nz(object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value);
			object returnItem = Invoker.MethodReturn(this, "Nz", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object LoadPicture(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "LoadPicture", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objtyp">Int32 objtyp</param>
		/// <param name="moduleName">string ModuleName</param>
		/// <param name="fileName">string FileName</param>
		/// <param name="token">Int32 token</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void ReplaceModule(Int32 objtyp, string moduleName, string fileName, Int32 token)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objtyp, moduleName, fileName, token);
			Invoker.Method(this, "ReplaceModule", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196488.aspx
		/// </summary>
		/// <param name="errorNumber">object ErrorNumber</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object AccessError(object errorNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(errorNumber);
			object returnItem = Invoker.MethodReturn(this, "AccessError", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object BuilderString()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BuilderString", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="guid">object Guid</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object StringFromGUID(object guid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(guid);
			object returnItem = Invoker.MethodReturn(this, "StringFromGUID", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="_string">object String</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object GUIDFromString(object _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(this, "GUIDFromString", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="id">Int32 id</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public object AppLoadString(Int32 id)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id);
			object returnItem = Invoker.MethodReturn(this, "AppLoadString", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822080.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="fileName">string FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SaveAsText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(this, "SaveAsText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="fileName">string FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void LoadFromText(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(this, "LoadFromText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823011.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToFavorites", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194960.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RefreshDatabaseWindow()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshDatabaseWindow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191909.aspx
		/// </summary>
		/// <param name="command">NetOffice.AccessApi.Enums.AcCommand Command</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void RunCommand(NetOffice.AccessApi.Enums.AcCommand command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(command);
			Invoker.Method(this, "RunCommand", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx
		/// </summary>
		/// <param name="hyperlink">object Hyperlink</param>
		/// <param name="part">optional NetOffice.AccessApi.Enums.AcHyperlinkPart Part = 0</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string HyperlinkPart(object hyperlink, object part)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hyperlink, part);
			object returnItem = Invoker.MethodReturn(this, "HyperlinkPart", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844740.aspx
		/// </summary>
		/// <param name="hyperlink">object Hyperlink</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public string HyperlinkPart(object hyperlink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hyperlink);
			object returnItem = Invoker.MethodReturn(this, "HyperlinkPart", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821756.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public bool GetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			object returnItem = Invoker.MethodReturn(this, "GetHiddenAttribute", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822459.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="fHidden">bool fHidden</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void SetHiddenAttribute(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, bool fHidden)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fHidden);
			Invoker.Method(this, "SetHiddenAttribute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		/// <param name="createNewFile">optional bool CreateNewFile = true</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName, object createNewFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, createNewFile);
			object returnItem = Invoker.MethodReturn(this, "CreateDataAccessPage", paramsArray);
			NetOffice.AccessApi.DataAccessPage newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.DataAccessPage CreateDataAccessPage()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateDataAccessPage", paramsArray);
			NetOffice.AccessApi.DataAccessPage newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.DataAccessPage CreateDataAccessPage(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "CreateDataAccessPage", paramsArray);
			NetOffice.AccessApi.DataAccessPage newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.DataAccessPage.LateBindingApiWrapperType) as NetOffice.AccessApi.DataAccessPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object Connect</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void NewAccessProject(string filepath, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, connect);
			Invoker.Method(this, "NewAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835758.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void NewAccessProject(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "NewAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenAccessProject(string filepath, object exclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive);
			Invoker.Method(this, "OpenAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837249.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void OpenAccessProject(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "OpenAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="connect">optional object Connect</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CreateAccessProject(string filepath, object connect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, connect);
			Invoker.Method(this, "CreateAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195216.aspx
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public void CreateAccessProject(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "CreateAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx
		/// </summary>
		/// <param name="number">Double Number</param>
		/// <param name="sourceCurrency">string SourceCurrency</param>
		/// <param name="targetCurrency">string TargetCurrency</param>
		/// <param name="fullPrecision">optional object FullPrecision</param>
		/// <param name="triangulationPrecision">optional object TriangulationPrecision</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision);
			object returnItem = Invoker.MethodReturn(this, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx
		/// </summary>
		/// <param name="number">Double Number</param>
		/// <param name="sourceCurrency">string SourceCurrency</param>
		/// <param name="targetCurrency">string TargetCurrency</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency);
			object returnItem = Invoker.MethodReturn(this, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192461.aspx
		/// </summary>
		/// <param name="number">Double Number</param>
		/// <param name="sourceCurrency">string SourceCurrency</param>
		/// <param name="targetCurrency">string TargetCurrency</param>
		/// <param name="fullPrecision">optional object FullPrecision</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14,15,16)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency, fullPrecision);
			object returnItem = Invoker.MethodReturn(this, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filepath">string filepath</param>
		/// <param name="exclusive">optional bool Exclusive = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenCurrentDatabaseOld(string filepath, object exclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath, exclusive);
			Invoker.Method(this, "OpenCurrentDatabaseOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void OpenCurrentDatabaseOld(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "OpenCurrentDatabaseOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		/// <param name="replace">optional bool Replace = false</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID, object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, company, workgroupID, replace);
			Invoker.Method(this, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CreateNewWorkgroupFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CreateNewWorkgroupFile(object path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CreateNewWorkgroupFile(object path, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name);
			Invoker.Method(this, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CreateNewWorkgroupFile(object path, object name, object company)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, company);
			Invoker.Method(this, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="path">optional string Path =  </param>
		/// <param name="name">optional string Name =  </param>
		/// <param name="company">optional string Company =  </param>
		/// <param name="workgroupID">optional string WorkgroupID =  </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void CreateNewWorkgroupFile(object path, object name, object company, object workgroupID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, name, company, workgroupID);
			Invoker.Method(this, "CreateNewWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195103.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void SetDefaultWorkgroupFile(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "SetDefaultWorkgroupFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193465.aspx
		/// </summary>
		/// <param name="sourceFilename">string SourceFilename</param>
		/// <param name="destinationFilename">string DestinationFilename</param>
		/// <param name="destinationFileFormat">NetOffice.AccessApi.Enums.AcFileFormat DestinationFileFormat</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ConvertAccessProject(string sourceFilename, string destinationFilename, NetOffice.AccessApi.Enums.AcFileFormat destinationFileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceFilename, destinationFilename, destinationFileFormat);
			Invoker.Method(this, "ConvertAccessProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx
		/// </summary>
		/// <param name="sourceFile">string SourceFile</param>
		/// <param name="destinationFile">string DestinationFile</param>
		/// <param name="logFile">optional bool LogFile = false</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool CompactRepair(string sourceFile, string destinationFile, object logFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceFile, destinationFile, logFile);
			object returnItem = Invoker.MethodReturn(this, "CompactRepair", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193242.aspx
		/// </summary>
		/// <param name="sourceFile">string SourceFile</param>
		/// <param name="destinationFile">string DestinationFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public bool CompactRepair(string sourceFile, string destinationFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sourceFile, destinationFile);
			object returnItem = Invoker.MethodReturn(this, "CompactRepair", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional Int32 OtherFlags = 0</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional NetOffice.AccessApi.Enums.AcExportXMLOtherFlags OtherFlags = 0</param>
		/// <param name="whereCondition">optional string WhereCondition = </param>
		/// <param name="additionalData">optional object AdditionalData</param>
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition, object additionalData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition, additionalData);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193212.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional NetOffice.AccessApi.Enums.AcExportXMLOtherFlags OtherFlags = 0</param>
		/// <param name="whereCondition">optional string WhereCondition = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXML(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags, object whereCondition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags, whereCondition);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx
		/// </summary>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="importOptions">optional NetOffice.AccessApi.Enums.AcImportXMLOption ImportOptions = 1</param>
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ImportXML(string dataSource, object importOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, importOptions);
			Invoker.Method(this, "ImportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823157.aspx
		/// </summary>
		/// <param name="dataSource">string DataSource</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 10,11,12,14,15,16)]
		public void ImportXML(string dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource);
			Invoker.Method(this, "ImportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		/// <param name="otherFlags">optional Int32 OtherFlags = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding, object otherFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding, otherFlags);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcExportXMLObjectType ObjectType</param>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="dataTarget">optional string DataTarget = </param>
		/// <param name="schemaTarget">optional string SchemaTarget = </param>
		/// <param name="presentationTarget">optional string PresentationTarget = </param>
		/// <param name="imageTarget">optional string ImageTarget = </param>
		/// <param name="encoding">optional NetOffice.AccessApi.Enums.AcExportXMLEncoding Encoding = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void ExportXMLOld(NetOffice.AccessApi.Enums.AcExportXMLObjectType objectType, string dataSource, object dataTarget, object schemaTarget, object presentationTarget, object imageTarget, object encoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, dataSource, dataTarget, schemaTarget, presentationTarget, imageTarget, encoding);
			Invoker.Method(this, "ExportXMLOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx
		/// </summary>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="transformSource">string TransformSource</param>
		/// <param name="outputTarget">string OutputTarget</param>
		/// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
		/// <param name="scriptOption">optional NetOffice.AccessApi.Enums.AcTransformXMLScriptOption ScriptOption = 1</param>
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput, object scriptOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, transformSource, outputTarget, wellFormedXMLOutput, scriptOption);
			Invoker.Method(this, "TransformXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx
		/// </summary>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="transformSource">string TransformSource</param>
		/// <param name="outputTarget">string OutputTarget</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void TransformXML(string dataSource, string transformSource, string outputTarget)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, transformSource, outputTarget);
			Invoker.Method(this, "TransformXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844810.aspx
		/// </summary>
		/// <param name="dataSource">string DataSource</param>
		/// <param name="transformSource">string TransformSource</param>
		/// <param name="outputTarget">string OutputTarget</param>
		/// <param name="wellFormedXMLOutput">optional bool WellFormedXMLOutput = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public void TransformXML(string dataSource, string transformSource, string outputTarget, object wellFormedXMLOutput)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataSource, transformSource, outputTarget, wellFormedXMLOutput);
			Invoker.Method(this, "TransformXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834773.aspx
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public NetOffice.AccessApi._AdditionalData CreateAdditionalData()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateAdditionalData", paramsArray);
			NetOffice.AccessApi._AdditionalData newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._AdditionalData;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dispid);
			object returnItem = Invoker.MethodReturn(this, "IsMemberSafe", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filepath">string filepath</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void NewCurrentDatabaseOld(string filepath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filepath);
			Invoker.Method(this, "NewCurrentDatabaseOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "CreateControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">optional NetOffice.AccessApi.Enums.AcSection Section = 0</param>
		/// <param name="parent">optional object Parent</param>
		/// <param name="columnName">optional object ColumnName</param>
		/// <param name="left">optional object Left</param>
		/// <param name="top">optional object Top</param>
		/// <param name="width">optional object Width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, object section, object parent, object columnName, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, columnName, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formName">string FormName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection Section</param>
		/// <param name="parent">string Parent</param>
		/// <param name="controlSource">string ControlSource</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateControlExOld(string formName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlSource, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formName, controlType, section, parent, controlSource, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateControlExOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="reportName">string ReportName</param>
		/// <param name="controlType">NetOffice.AccessApi.Enums.AcControlType ControlType</param>
		/// <param name="section">NetOffice.AccessApi.Enums.AcSection Section</param>
		/// <param name="parent">string Parent</param>
		/// <param name="controlName">string ControlName</param>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public NetOffice.AccessApi.Control CreateReportControlExOld(string reportName, NetOffice.AccessApi.Enums.AcControlType controlType, NetOffice.AccessApi.Enums.AcSection section, string parent, string controlName, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(reportName, controlType, section, parent, controlName, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "CreateReportControlExOld", paramsArray);
			NetOffice.AccessApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.AccessApi.Control.LateBindingApiWrapperType) as NetOffice.AccessApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx
		/// </summary>
		/// <param name="richText">object RichText</param>
		/// <param name="length">optional object Length</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string PlainText(object richText, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(richText, length);
			object returnItem = Invoker.MethodReturn(this, "PlainText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196012.aspx
		/// </summary>
		/// <param name="richText">object RichText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string PlainText(object richText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(richText);
			object returnItem = Invoker.MethodReturn(this, "PlainText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx
		/// </summary>
		/// <param name="plainText">object PlainText</param>
		/// <param name="length">optional object Length</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string HtmlEncode(object plainText, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(plainText, length);
			object returnItem = Invoker.MethodReturn(this, "HtmlEncode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192028.aspx
		/// </summary>
		/// <param name="plainText">object PlainText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string HtmlEncode(object plainText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(plainText);
			object returnItem = Invoker.MethodReturn(this, "HtmlEncode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194416.aspx
		/// </summary>
		/// <param name="customUIName">string CustomUIName</param>
		/// <param name="customUIXML">string CustomUIXML</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void LoadCustomUI(string customUIName, string customUIXML)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customUIName, customUIXML);
			Invoker.Method(this, "LoadCustomUI", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193467.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ExportNavigationPane(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "ExportNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fAppendOnly">optional bool fAppendOnly = false</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ImportNavigationPane(string path, object fAppendOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fAppendOnly);
			Invoker.Method(this, "ImportNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193985.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ImportNavigationPane(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "ImportNavigationPane", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835727.aspx
		/// </summary>
		/// <param name="tableName">string TableName</param>
		/// <param name="columnName">string ColumnName</param>
		/// <param name="queryString">string queryString</param>
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public string ColumnHistory(string tableName, string columnName, string queryString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tableName, columnName, queryString);
			object returnItem = Invoker.MethodReturn(this, "ColumnHistory", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="externalExporter">object ExternalExporter</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage, object toPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage, toPage);
			Invoker.Method(this, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="externalExporter">object ExternalExporter</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType);
			Invoker.Method(this, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="externalExporter">object ExternalExporter</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType, selectedRecords);
			Invoker.Method(this, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="externalExporter">object ExternalExporter</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcOutputObjectType ObjectType</param>
		/// <param name="selectedRecords">optional bool SelectedRecords = false</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 12,14,15,16)]
		public void ExportCustomFixedFormat(object externalExporter, string outputFileName, string objectName, NetOffice.AccessApi.Enums.AcOutputObjectType objectType, object selectedRecords, object fromPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(externalExporter, outputFileName, objectName, objectType, selectedRecords, fromPage);
			Invoker.Method(this, "ExportCustomFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821429.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(this, "SaveAsAXL", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845765.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void LoadFromAXL(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName, fileName);
			Invoker.Method(this, "LoadFromAXL", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		/// <param name="previewPath">optional object PreviewPath</param>
		/// <param name="description">optional object Description</param>
		/// <param name="instantiationForm">optional object InstantiationForm</param>
		/// <param name="applicationPart">optional object ApplicationPart</param>
		/// <param name="includeData">optional object IncludeData</param>
		/// <param name="variation">optional object Variation</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData, object variation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData, variation);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		/// <param name="previewPath">optional object PreviewPath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		/// <param name="previewPath">optional object PreviewPath</param>
		/// <param name="description">optional object Description</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		/// <param name="previewPath">optional object PreviewPath</param>
		/// <param name="description">optional object Description</param>
		/// <param name="instantiationForm">optional object InstantiationForm</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		/// <param name="previewPath">optional object PreviewPath</param>
		/// <param name="description">optional object Description</param>
		/// <param name="instantiationForm">optional object InstantiationForm</param>
		/// <param name="applicationPart">optional object ApplicationPart</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192852.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="title">string Title</param>
		/// <param name="iconPath">string IconPath</param>
		/// <param name="coreTable">string CoreTable</param>
		/// <param name="category">string Category</param>
		/// <param name="previewPath">optional object PreviewPath</param>
		/// <param name="description">optional object Description</param>
		/// <param name="instantiationForm">optional object InstantiationForm</param>
		/// <param name="applicationPart">optional object ApplicationPart</param>
		/// <param name="includeData">optional object IncludeData</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void SaveAsTemplate(string path, string title, string iconPath, string coreTable, string category, object previewPath, object description, object instantiationForm, object applicationPart, object includeData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, title, iconPath, coreTable, category, previewPath, description, instantiationForm, applicationPart, includeData);
			Invoker.Method(this, "SaveAsTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835421.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void InstantiateTemplate(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "InstantiateTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834388.aspx
		/// </summary>
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserDisplay DisplayOption</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public object CurrentWebUser(NetOffice.AccessApi.Enums.AcWebUserDisplay displayOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(displayOption);
			object returnItem = Invoker.MethodReturn(this, "CurrentWebUser", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="displayOption">NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay DisplayOption</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public object CurrentWebUserGroups(NetOffice.AccessApi.Enums.AcWebUserGroupsDisplay displayOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(displayOption);
			object returnItem = Invoker.MethodReturn(this, "CurrentWebUserGroups", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		/// <param name="groupNameOrID">object GroupNameOrID</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public bool IsCurrentWebUserInGroup(object groupNameOrID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupNameOrID);
			object returnItem = Invoker.MethodReturn(this, "IsCurrentWebUserInGroup", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834368.aspx
		/// </summary>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType ObjectType</param>
		/// <param name="objectName">string ObjectName</param>
		[SupportByVersionAttribute("Access", 14,15,16)]
		public void DirtyObject(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectType, objectName);
			Invoker.Method(this, "DirtyObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 14,15,16)]
		public bool IsClient()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "IsClient", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}