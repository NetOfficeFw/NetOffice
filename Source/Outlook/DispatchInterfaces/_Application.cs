using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// DispatchInterface _Application 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868973.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OutlookApi._Application newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865581.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Class", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlObjectClass)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866436.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Session", paramsArray);
				NetOffice.OutlookApi._NameSpace newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NameSpace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869381.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868248.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860684.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870066.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868795.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Explorers Explorers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Explorers", paramsArray);
				NetOffice.OutlookApi._Explorers newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Explorers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868935.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Inspectors Inspectors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Inspectors", paramsArray);
				NetOffice.OutlookApi._Inspectors newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Inspectors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867217.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869152.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862144.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Reminders Reminders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Reminders", paramsArray);
				NetOffice.OutlookApi._Reminders newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Reminders;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865059.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public string DefaultProfileName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultProfileName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864729.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public bool IsTrusted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsTrusted", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861554.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff867696.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.TimeZones TimeZones
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TimeZones", paramsArray);
				NetOffice.OutlookApi.TimeZones newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OutlookApi.TimeZones.LateBindingApiWrapperType) as NetOffice.OutlookApi.TimeZones;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861549.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public NetOffice.OfficeApi.PickerDialog PickerDialog
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PickerDialog", paramsArray);
				NetOffice.OfficeApi.PickerDialog newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.PickerDialog.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerDialog;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff870017.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Explorer ActiveExplorer()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ActiveExplorer", paramsArray);
			NetOffice.OutlookApi._Explorer newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Explorer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863939.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._Inspector ActiveInspector()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ActiveInspector", paramsArray);
			NetOffice.OutlookApi._Inspector newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Inspector;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869635.aspx
		/// </summary>
		/// <param name="itemType">NetOffice.OutlookApi.Enums.OlItemType ItemType</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object CreateItem(NetOffice.OutlookApi.Enums.OlItemType itemType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(itemType);
			object returnItem = Invoker.MethodReturn(this, "CreateItem", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865637.aspx
		/// </summary>
		/// <param name="templatePath">string TemplatePath</param>
		/// <param name="inFolder">optional object InFolder</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object CreateItemFromTemplate(string templatePath, object inFolder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(templatePath, inFolder);
			object returnItem = Invoker.MethodReturn(this, "CreateItemFromTemplate", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865637.aspx
		/// </summary>
		/// <param name="templatePath">string TemplatePath</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object CreateItemFromTemplate(string templatePath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(templatePath);
			object returnItem = Invoker.MethodReturn(this, "CreateItemFromTemplate", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860724.aspx
		/// </summary>
		/// <param name="objectName">string ObjectName</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object CreateObject(string objectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectName);
			object returnItem = Invoker.MethodReturn(this, "CreateObject", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865800.aspx
		/// </summary>
		/// <param name="type">string Type</param>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi._NameSpace GetNamespace(string type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "GetNamespace", paramsArray);
			NetOffice.OutlookApi._NameSpace newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NameSpace;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866010.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865654.aspx
		/// </summary>
		[SupportByVersionAttribute("Outlook", 9,10,11,12,14,15,16)]
		public object ActiveWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ActiveWindow", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869462.aspx
		/// </summary>
		/// <param name="filePath">string FilePath</param>
		/// <param name="destFolderPath">string DestFolderPath</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public object CopyFile(string filePath, string destFolderPath)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filePath, destFolderPath);
			object returnItem = Invoker.MethodReturn(this, "CopyFile", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx
		/// </summary>
		/// <param name="scope">string Scope</param>
		/// <param name="filter">optional object Filter</param>
		/// <param name="searchSubFolders">optional object SearchSubFolders</param>
		/// <param name="tag">optional object Tag</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders, object tag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scope, filter, searchSubFolders, tag);
			object returnItem = Invoker.MethodReturn(this, "AdvancedSearch", paramsArray);
			NetOffice.OutlookApi.Search newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Search.LateBindingApiWrapperType) as NetOffice.OutlookApi.Search;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx
		/// </summary>
		/// <param name="scope">string Scope</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scope);
			object returnItem = Invoker.MethodReturn(this, "AdvancedSearch", paramsArray);
			NetOffice.OutlookApi.Search newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Search.LateBindingApiWrapperType) as NetOffice.OutlookApi.Search;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx
		/// </summary>
		/// <param name="scope">string Scope</param>
		/// <param name="filter">optional object Filter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scope, filter);
			object returnItem = Invoker.MethodReturn(this, "AdvancedSearch", paramsArray);
			NetOffice.OutlookApi.Search newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Search.LateBindingApiWrapperType) as NetOffice.OutlookApi.Search;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866933.aspx
		/// </summary>
		/// <param name="scope">string Scope</param>
		/// <param name="filter">optional object Filter</param>
		/// <param name="searchSubFolders">optional object SearchSubFolders</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Search AdvancedSearch(string scope, object filter, object searchSubFolders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scope, filter, searchSubFolders);
			object returnItem = Invoker.MethodReturn(this, "AdvancedSearch", paramsArray);
			NetOffice.OutlookApi.Search newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OutlookApi.Search.LateBindingApiWrapperType) as NetOffice.OutlookApi.Search;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869145.aspx
		/// </summary>
		/// <param name="lookInFolders">string LookInFolders</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public bool IsSearchSynchronous(string lookInFolders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lookInFolders);
			object returnItem = Invoker.MethodReturn(this, "IsSearchSynchronous", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvar">object pvar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Outlook", 10,11,12,14,15,16)]
		public void GetNewNickNames(object pvar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvar);
			Invoker.Method(this, "GetNewNickNames", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864203.aspx
		/// </summary>
		/// <param name="item">object Item</param>
		/// <param name="referenceType">NetOffice.OutlookApi.Enums.OlReferenceType ReferenceType</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public object GetObjectReference(object item, NetOffice.OutlookApi.Enums.OlReferenceType referenceType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(item, referenceType);
			object returnItem = Invoker.MethodReturn(this, "GetObjectReference", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863374.aspx
		/// </summary>
		/// <param name="regionName">string RegionName</param>
		[SupportByVersionAttribute("Outlook", 14,15,16)]
		public void RefreshFormRegionDefinition(string regionName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(regionName);
			Invoker.Method(this, "RefreshFormRegionDefinition", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}