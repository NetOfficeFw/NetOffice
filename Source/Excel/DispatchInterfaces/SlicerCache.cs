using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface SlicerCache 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822652.aspx </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SlicerCache : COMObject
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
                    _type = typeof(SlicerCache);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SlicerCache(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SlicerCache(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlicerCache(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlicerCache(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlicerCache(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlicerCache(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlicerCache() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlicerCache(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837149.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821255.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834306.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838257.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821819.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool OLAP
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OLAP");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198150.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPivotTableSourceType SourceType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotTableSourceType>(this, "SourceType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841279.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.WorkbookConnection WorkbookConnection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookConnection>(this, "WorkbookConnection", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836510.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicers Slicers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicers>(this, "Slicers", NetOffice.ExcelApi.Slicers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823056.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerPivotTables PivotTables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerPivotTables>(this, "PivotTables", NetOffice.ExcelApi.SlicerPivotTables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193891.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerCacheLevels SlicerCacheLevels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerCacheLevels>(this, "SlicerCacheLevels", NetOffice.ExcelApi.SlicerCacheLevels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196889.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840475.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerItems VisibleSlicerItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerItems>(this, "VisibleSlicerItems", NetOffice.ExcelApi.SlicerItems.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193916.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public object VisibleSlicerItemsList
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "VisibleSlicerItemsList");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "VisibleSlicerItemsList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839561.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerItems SlicerItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerItems>(this, "SlicerItems", NetOffice.ExcelApi.SlicerItems.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835315.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSlicerCrossFilterType CrossFilterType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSlicerCrossFilterType>(this, "CrossFilterType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CrossFilterType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839813.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSlicerSort SortItems
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSlicerSort>(this, "SortItems");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SortItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821964.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public string SourceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821809.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool SortUsingCustomLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SortUsingCustomLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SortUsingCustomLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822904.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool ShowAllItems
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAllItems");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAllItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232131.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.TimelineState TimelineState
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TimelineState>(this, "TimelineState", NetOffice.ExcelApi.TimelineState.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231161.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlSlicerCacheType SlicerCacheType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSlicerCacheType>(this, "SlicerCacheType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230306.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool FilterCleared
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterCleared");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229894.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool List
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "List");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229518.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool RequireManualUpdate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RequireManualUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RequireManualUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230733.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ListObject ListObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObject>(this, "ListObject", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229786.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void ClearManualFilter()
		{
			 Factory.ExecuteMethod(this, "ClearManualFilter");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196378.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229270.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public void ClearAllFilters()
		{
			 Factory.ExecuteMethod(this, "ClearAllFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231775.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public void ClearDateFilter()
		{
			 Factory.ExecuteMethod(this, "ClearDateFilter");
		}

		#endregion

		#pragma warning restore
	}
}
