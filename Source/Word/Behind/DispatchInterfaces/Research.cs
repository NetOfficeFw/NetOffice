using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Research 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194717.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Research : COMObject, NetOffice.WordApi.Research
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.WordApi.Research);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Research() : base()
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
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196335.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845563.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840115.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual string FavoriteService
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FavoriteService");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FavoriteService", value);
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
		public virtual object Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Query", new object[]{ serviceID, queryString, queryLanguage, useSelection, launchQuery });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual object Query(string serviceID)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Query", serviceID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194387.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		/// <param name="queryString">optional string QueryString = </param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual object Query(string serviceID, object queryString)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Query", serviceID, queryString);
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
		public virtual object Query(string serviceID, object queryString, object queryLanguage)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Query", serviceID, queryString, queryLanguage);
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
		public virtual object Query(string serviceID, object queryString, object queryLanguage, object useSelection)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Query", serviceID, queryString, queryLanguage, useSelection);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834572.aspx </remarks>
		/// <param name="languageFrom">NetOffice.WordApi.Enums.WdLanguageID languageFrom</param>
		/// <param name="languageTo">NetOffice.WordApi.Enums.WdLanguageID languageTo</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual object SetLanguagePair(NetOffice.WordApi.Enums.WdLanguageID languageFrom, NetOffice.WordApi.Enums.WdLanguageID languageTo)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SetLanguagePair", languageFrom, languageTo);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835810.aspx </remarks>
		/// <param name="serviceID">string serviceID</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool IsResearchService(string serviceID)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsResearchService", serviceID);
		}

		#endregion

		#pragma warning restore
	}
}


