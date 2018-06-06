using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface SlicerCache 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822652.aspx </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SlicerCache : COMObject, NetOffice.ExcelApi.SlicerCache
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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SlicerCache() : base()
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
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821255.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
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
		public virtual object Parent
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
		public virtual Int32 Index
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
		public virtual bool OLAP
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
		public virtual NetOffice.ExcelApi.Enums.XlPivotTableSourceType SourceType
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
		public virtual NetOffice.ExcelApi.WorkbookConnection WorkbookConnection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookConnection>(this, "WorkbookConnection", typeof(NetOffice.ExcelApi.WorkbookConnection));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836510.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Slicers Slicers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicers>(this, "Slicers", typeof(NetOffice.ExcelApi.Slicers));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823056.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.SlicerPivotTables PivotTables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerPivotTables>(this, "PivotTables", typeof(NetOffice.ExcelApi.SlicerPivotTables));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193891.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.SlicerCacheLevels SlicerCacheLevels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerCacheLevels>(this, "SlicerCacheLevels", typeof(NetOffice.ExcelApi.SlicerCacheLevels));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196889.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual string Name
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
		public virtual NetOffice.ExcelApi.SlicerItems VisibleSlicerItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerItems>(this, "VisibleSlicerItems", typeof(NetOffice.ExcelApi.SlicerItems));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193916.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual object VisibleSlicerItemsList
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
		public virtual NetOffice.ExcelApi.SlicerItems SlicerItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerItems>(this, "SlicerItems", typeof(NetOffice.ExcelApi.SlicerItems));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835315.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlSlicerCrossFilterType CrossFilterType
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
		public virtual NetOffice.ExcelApi.Enums.XlSlicerSort SortItems
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
		public virtual string SourceName
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
		public virtual bool SortUsingCustomLists
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
		public virtual bool ShowAllItems
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
		public virtual NetOffice.ExcelApi.TimelineState TimelineState
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TimelineState>(this, "TimelineState", typeof(NetOffice.ExcelApi.TimelineState));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231161.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Enums.XlSlicerCacheType SlicerCacheType
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
		public virtual bool FilterCleared
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
		public virtual bool List
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
		public virtual bool RequireManualUpdate
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
		public virtual NetOffice.ExcelApi.ListObject ListObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObject>(this, "ListObject", typeof(NetOffice.ExcelApi.ListObject));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229786.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual void ClearManualFilter()
		{
			 Factory.ExecuteMethod(this, "ClearManualFilter");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196378.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229270.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual void ClearAllFilters()
		{
			 Factory.ExecuteMethod(this, "ClearAllFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231775.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual void ClearDateFilter()
		{
			 Factory.ExecuteMethod(this, "ClearDateFilter");
		}

		#endregion

		#pragma warning restore
	}
}


