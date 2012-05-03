using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface System 
	/// SupportByVersion Word, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class System : COMObject
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
                    _type = typeof(System);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public System(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public System(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public System(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public System() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public System(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string OperatingSystem
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OperatingSystem", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string ProcessorType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProcessorType", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 FreeDiskSpace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FreeDiskSpace", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Enums.WdCountry Country
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Country", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdCountry)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string LanguageDesignation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageDesignation", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 HorizontalResolution
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HorizontalResolution", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 VerticalResolution
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VerticalResolution", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		/// <param name="section">string Section</param>
		/// <param name="key">string Key</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ProfileString(string section, string key)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, key);
			object returnItem = Invoker.PropertyGet(this, "ProfileString", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		/// <param name="section">string Section</param>
		/// <param name="key">string Key</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void set_ProfileString(string section, string key, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section, key);
			Invoker.PropertySet(this, "ProfileString", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Alias for get_ProfileString
		/// </summary>
		/// <param name="section">string Section</param>
		/// <param name="key">string Key</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string ProfileString(string section, string key)
		{
			return get_ProfileString(section, key);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">string Section</param>
		/// <param name="key">string Key</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_PrivateProfileString(string fileName, string section, string key)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section, key);
			object returnItem = Invoker.PropertyGet(this, "PrivateProfileString", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">string Section</param>
		/// <param name="key">string Key</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void set_PrivateProfileString(string fileName, string section, string key, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, section, key);
			Invoker.PropertySet(this, "PrivateProfileString", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Alias for get_PrivateProfileString
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="section">string Section</param>
		/// <param name="key">string Key</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string PrivateProfileString(string fileName, string section, string key)
		{
			return get_PrivateProfileString(fileName, section, key);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public bool MathCoprocessorInstalled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MathCoprocessorInstalled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string ComputerType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ComputerType", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string MacintoshName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MacintoshName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public bool QuickDrawInstalled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QuickDrawInstalled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Enums.WdCursorType Cursor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cursor", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdCursorType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Cursor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public NetOffice.WordApi.Enums.WdCountry CountryRegion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CountryRegion", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdCountry)intReturnItem;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void MSInfo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MSInfo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="drive">optional object Drive</param>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Connect(string path, object drive, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, drive, password);
			Invoker.Method(this, "Connect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Connect(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "Connect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="drive">optional object Drive</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Connect(string path, object drive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, drive);
			Invoker.Method(this, "Connect", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}