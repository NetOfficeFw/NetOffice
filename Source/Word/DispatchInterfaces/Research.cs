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
	/// DispatchInterface Research 
	/// SupportByVersion Word, 12,14
	///</summary>
	[SupportByVersionAttribute("Word", 12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Research : COMObject
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
                    _type = typeof(Research);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public string FavoriteService
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FavoriteService", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FavoriteService", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		/// <param name="launchQuery">optional bool LaunchQuery = true</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public object Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString, queryLanguage, useSelection, launchQuery);
			object returnItem = Invoker.MethodReturn(this, "Query", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public object Query(string serviceID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID);
			object returnItem = Invoker.MethodReturn(this, "Query", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public object Query(string serviceID, object queryString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString);
			object returnItem = Invoker.MethodReturn(this, "Query", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public object Query(string serviceID, object queryString, object queryLanguage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString, queryLanguage);
			object returnItem = Invoker.MethodReturn(this, "Query", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public object Query(string serviceID, object queryString, object queryLanguage, object useSelection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString, queryLanguage, useSelection);
			object returnItem = Invoker.MethodReturn(this, "Query", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="languageFrom">NetOffice.WordApi.Enums.WdLanguageID LanguageFrom</param>
		/// <param name="languageTo">NetOffice.WordApi.Enums.WdLanguageID LanguageTo</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public object SetLanguagePair(NetOffice.WordApi.Enums.WdLanguageID languageFrom, NetOffice.WordApi.Enums.WdLanguageID languageTo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(languageFrom, languageTo);
			object returnItem = Invoker.MethodReturn(this, "SetLanguagePair", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool IsResearchService(string serviceID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID);
			object returnItem = Invoker.MethodReturn(this, "IsResearchService", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}