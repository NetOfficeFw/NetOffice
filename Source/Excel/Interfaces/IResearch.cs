using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IResearch 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IResearch : COMObject
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
                    _type = typeof(IResearch);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IResearch(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IResearch(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IResearch(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IResearch(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IResearch(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IResearch(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IResearch() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IResearch(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		/// <param name="useSelection">optional object useSelection</param>
		/// <param name="launchQuery">optional object launchQuery</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", new object[]{ serviceID, queryString, queryLanguage, useSelection, launchQuery });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Query(string serviceID)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Query(string serviceID, object queryString)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID, queryString);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Query(string serviceID, object queryString, object queryLanguage)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID, queryString, queryLanguage);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional object queryString</param>
		/// <param name="queryLanguage">optional object queryLanguage</param>
		/// <param name="useSelection">optional object useSelection</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Query(string serviceID, object queryString, object queryLanguage, object useSelection)
		{
			return Factory.ExecuteVariantMethodGet(this, "Query", serviceID, queryString, queryLanguage, useSelection);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool IsResearchService(string serviceID)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsResearchService", serviceID);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="languageFrom">Int32 languageFrom</param>
		/// <param name="languageTo">Int32 languageTo</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object SetLanguagePair(Int32 languageFrom, Int32 languageTo)
		{
			return Factory.ExecuteVariantMethodGet(this, "SetLanguagePair", languageFrom, languageTo);
		}

		#endregion

		#pragma warning restore
	}
}
