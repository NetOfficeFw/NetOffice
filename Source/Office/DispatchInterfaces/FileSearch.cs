using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface FileSearch 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class FileSearch : _IMsoDispObj
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
                    _type = typeof(FileSearch);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public FileSearch(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FileSearch(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileSearch(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileSearch(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileSearch(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileSearch(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileSearch() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileSearch(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool SearchSubFolders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SearchSubFolders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SearchSubFolders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool MatchTextExactly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchTextExactly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchTextExactly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool MatchAllWordForms
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchAllWordForms");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchAllWordForms", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string FileName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FileName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FileName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFileType FileType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileType>(this, "FileType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FileType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLastModified LastModified
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLastModified>(this, "LastModified");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LastModified", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string TextOrProperty
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextOrProperty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextOrProperty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public string LookIn
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LookIn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LookIn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FoundFiles FoundFiles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FoundFiles>(this, "FoundFiles", NetOffice.OfficeApi.FoundFiles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.PropertyTests PropertyTests
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PropertyTests>(this, "PropertyTests", NetOffice.OfficeApi.PropertyTests.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SearchScopes SearchScopes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SearchScopes>(this, "SearchScopes", NetOffice.OfficeApi.SearchScopes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SearchFolders SearchFolders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SearchFolders>(this, "SearchFolders", NetOffice.OfficeApi.SearchFolders.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileTypes FileTypes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileTypes>(this, "FileTypes", NetOffice.OfficeApi.FileTypes.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
		/// <param name="sortOrder">optional NetOffice.OfficeApi.Enums.MsoSortOrder SortOrder = 1</param>
		/// <param name="alwaysAccurate">optional bool AlwaysAccurate = true</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Execute(object sortBy, object sortOrder, object alwaysAccurate)
		{
			return Factory.ExecuteInt32MethodGet(this, "Execute", sortBy, sortOrder, alwaysAccurate);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Execute()
		{
			return Factory.ExecuteInt32MethodGet(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Execute(object sortBy)
		{
			return Factory.ExecuteInt32MethodGet(this, "Execute", sortBy);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
		/// <param name="sortOrder">optional NetOffice.OfficeApi.Enums.MsoSortOrder SortOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Execute(object sortBy, object sortOrder)
		{
			return Factory.ExecuteInt32MethodGet(this, "Execute", sortBy, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void NewSearch()
		{
			 Factory.ExecuteMethod(this, "NewSearch");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void RefreshScopes()
		{
			 Factory.ExecuteMethod(this, "RefreshScopes");
		}

		#endregion

		#pragma warning restore
	}
}
