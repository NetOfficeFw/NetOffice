using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Research 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194717.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Research : COMObject
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
                    _type = typeof(Research);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Research(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Research(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192412.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196335.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845563.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840115.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public string FavoriteService
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FavoriteService");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FavoriteService", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		/// <param name="launchQuery">optional bool LaunchQuery = true</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public object Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", new object[]{ serviceID, queryString, queryLanguage, useSelection, launchQuery });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public object Query(string serviceID)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public object Query(string serviceID, object queryString)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID, queryString);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public object Query(string serviceID, object queryString, object queryLanguage)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID, queryString, queryLanguage);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		/// <param name="queryLanguage">optional NetOffice.WordApi.Enums.WdLanguageID QueryLanguage = 0</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public object Query(string serviceID, object queryString, object queryLanguage, object useSelection)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID, queryString, queryLanguage, useSelection);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834572.aspx </remarks>
		/// <param name="languageFrom">NetOffice.WordApi.Enums.WdLanguageID languageFrom</param>
		/// <param name="languageTo">NetOffice.WordApi.Enums.WdLanguageID languageTo</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public object SetLanguagePair(NetOffice.WordApi.Enums.WdLanguageID languageFrom, NetOffice.WordApi.Enums.WdLanguageID languageTo)
		{
			return Factory.ExecuteVariantMethodGet(this, "SetLanguagePair", languageFrom, languageTo);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835810.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool IsResearchService(string serviceID)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsResearchService", serviceID);
		}

		#endregion

		#pragma warning restore
	}
}
