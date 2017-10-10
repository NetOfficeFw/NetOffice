using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// IPivotTable
	/// </summary>
	[SyntaxBypass]
 	public class IPivotTable_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IPivotTable_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IPivotTable_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ColumnFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "ColumnFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ColumnFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_ColumnFields")]
		public object ColumnFields(object index)
		{
			return get_ColumnFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_DataFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "DataFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DataFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_DataFields")]
		public object DataFields(object index)
		{
			return get_DataFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HiddenFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "HiddenFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_HiddenFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_HiddenFields")]
		public object HiddenFields(object index)
		{
			return get_HiddenFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_PageFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "PageFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_PageFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_PageFields")]
		public object PageFields(object index)
		{
			return get_PageFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_RowFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "RowFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_RowFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_RowFields")]
		public object RowFields(object index)
		{
			return get_RowFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_VisibleFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "VisibleFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_VisibleFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_VisibleFields")]
		public object VisibleFields(object index)
		{
			return get_VisibleFields(index);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// Interface IPivotTable 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IPivotTable : COMObject
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
                    _type = typeof(IPivotTable);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IPivotTable(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IPivotTable(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTable(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object ColumnFields
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ColumnFields");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ColumnGrand
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ColumnGrand");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnGrand", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range ColumnRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "ColumnRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DataBodyRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataBodyRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object DataFields
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DataFields");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DataLabelRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataLabelRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_Default", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool HasAutoFormat
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasAutoFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasAutoFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object HiddenFields
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "HiddenFields");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string InnerDetail
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InnerDetail");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InnerDetail", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object PageFields
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "PageFields");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range PageRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "PageRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range PageRangeCells
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "PageRangeCells", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public DateTime RefreshDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "RefreshDate");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string RefreshName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RefreshName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object RowFields
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "RowFields");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RowGrand
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RowGrand");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowGrand", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range RowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "RowRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SaveData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SourceData
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SourceData");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SourceData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range TableRange1
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "TableRange1", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range TableRange2
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "TableRange2", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object VisibleFields
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "VisibleFields");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CacheIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CacheIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CacheIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayErrorString
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayErrorString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayErrorString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayNullString
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayNullString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayNullString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableDrilldown
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableDrilldown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableDrilldown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableFieldDialog
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableFieldDialog");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableFieldDialog", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableWizard
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableWizard");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableWizard", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string ErrorString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ErrorString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ErrorString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ManualUpdate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ManualUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ManualUpdate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool MergeLabels
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MergeLabels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MergeLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string NullString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NullString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NullString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotFormulas PivotFormulas
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFormulas>(this, "PivotFormulas", NetOffice.ExcelApi.PivotFormulas.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SubtotalHiddenPageItems
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SubtotalHiddenPageItems");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubtotalHiddenPageItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PageFieldOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageFieldOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageFieldOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string PageFieldStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PageFieldStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageFieldStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PageFieldWrapCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageFieldWrapCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageFieldWrapCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PreserveFormatting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PreserveFormatting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PreserveFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string PivotSelection
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PivotSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PivotSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPTSelectionMode SelectionMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPTSelectionMode>(this, "SelectionMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SelectionMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string TableStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TableStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TableStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Tag
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string VacatedStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "VacatedStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VacatedStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PrintTitles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintTitles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintTitles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CubeFields CubeFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CubeFields>(this, "CubeFields", NetOffice.ExcelApi.CubeFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string GrandTotalName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "GrandTotalName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GrandTotalName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SmallGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SmallGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmallGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RepeatItemsOnEachPrintedPage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RepeatItemsOnEachPrintedPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RepeatItemsOnEachPrintedPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TotalsAnnotation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TotalsAnnotation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TotalsAnnotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string PivotSelectionStandard
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PivotSelectionStandard");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PivotSelectionStandard", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField DataPivotField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "DataPivotField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EnableDataValueEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableDataValueEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableDataValueEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string MDX
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MDX");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool ViewCalculatedMembers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewCalculatedMembers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewCalculatedMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedMembers CalculatedMembers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CalculatedMembers>(this, "CalculatedMembers", NetOffice.ExcelApi.CalculatedMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool DisplayImmediateItems
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayImmediateItems");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayImmediateItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EnableFieldList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableFieldList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool VisualTotals
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VisualTotals");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VisualTotals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool ShowPageMultipleItemLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPageMultipleItemLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPageMultipleItemLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPivotTableVersionList Version
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotTableVersionList>(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool DisplayEmptyRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayEmptyRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayEmptyRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool DisplayEmptyColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayEmptyColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayEmptyColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool ShowCellBackgroundFromOLAP
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowCellBackgroundFromOLAP");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowCellBackgroundFromOLAP", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotAxis PivotColumnAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotAxis>(this, "PivotColumnAxis", NetOffice.ExcelApi.PivotAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotAxis PivotRowAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotAxis>(this, "PivotRowAxis", NetOffice.ExcelApi.PivotAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowDrillIndicators
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDrillIndicators");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDrillIndicators", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool PrintDrillIndicators
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintDrillIndicators");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintDrillIndicators", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DisplayMemberPropertyTooltips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayMemberPropertyTooltips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayMemberPropertyTooltips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DisplayContextTooltips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayContextTooltips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayContextTooltips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 CompactRowIndent
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CompactRowIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CompactRowIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlLayoutRowType LayoutRowDefault
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlLayoutRowType>(this, "LayoutRowDefault");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutRowDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DisplayFieldCaptions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayFieldCaptions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayFieldCaptions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilters ActiveFilters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFilters>(this, "ActiveFilters", NetOffice.ExcelApi.PivotFilters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool InGridDropZones
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InGridDropZones");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InGridDropZones", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object TableStyle2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TableStyle2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TableStyle2", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleLastColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleLastColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleLastColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleRowStripes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleRowStripes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleRowStripes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleColumnStripes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleColumnStripes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleColumnStripes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleRowHeaders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleRowHeaders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleRowHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleColumnHeaders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleColumnHeaders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleColumnHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool AllowMultipleFilters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowMultipleFilters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowMultipleFilters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string CompactLayoutRowHeader
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompactLayoutRowHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CompactLayoutRowHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string CompactLayoutColumnHeader
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompactLayoutColumnHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CompactLayoutColumnHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool FieldListSortAscending
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FieldListSortAscending");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FieldListSortAscending", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string Location
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Location");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Location", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool EnableWriteback
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableWriteback");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableWriteback", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlAllocation Allocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAllocation>(this, "Allocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Allocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlAllocationValue AllocationValue
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAllocationValue>(this, "AllocationValue");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AllocationValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlAllocationMethod AllocationMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAllocationMethod>(this, "AllocationMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AllocationMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string AllocationWeightExpression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AllocationWeightExpression");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllocationWeightExpression", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.PivotTableChangeList ChangeList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotTableChangeList>(this, "ChangeList", NetOffice.ExcelApi.PivotTableChangeList.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string AlternativeText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AlternativeText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternativeText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string Summary
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Summary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Summary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool VisualTotalsForSets
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VisualTotalsForSets");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VisualTotalsForSets", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool ShowValuesRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowValuesRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowValuesRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool CalculatedMembersInFilters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CalculatedMembersInFilters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CalculatedMembersInFilters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		public bool Hidden
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Hidden");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape PivotChart
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shape>(this, "PivotChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="rowFields">optional object rowFields</param>
		/// <param name="columnFields">optional object columnFields</param>
		/// <param name="pageFields">optional object pageFields</param>
		/// <param name="addToTable">optional object addToTable</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddFields(object rowFields, object columnFields, object pageFields, object addToTable)
		{
			return Factory.ExecuteVariantMethodGet(this, "AddFields", rowFields, columnFields, pageFields, addToTable);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddFields()
		{
			return Factory.ExecuteVariantMethodGet(this, "AddFields");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="rowFields">optional object rowFields</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddFields(object rowFields)
		{
			return Factory.ExecuteVariantMethodGet(this, "AddFields", rowFields);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="rowFields">optional object rowFields</param>
		/// <param name="columnFields">optional object columnFields</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddFields(object rowFields, object columnFields)
		{
			return Factory.ExecuteVariantMethodGet(this, "AddFields", rowFields, columnFields);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="rowFields">optional object rowFields</param>
		/// <param name="columnFields">optional object columnFields</param>
		/// <param name="pageFields">optional object pageFields</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddFields(object rowFields, object columnFields, object pageFields)
		{
			return Factory.ExecuteVariantMethodGet(this, "AddFields", rowFields, columnFields, pageFields);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pageField">optional object pageField</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowPages(object pageField)
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowPages", pageField);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowPages()
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowPages");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotFields(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotFields()
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotFields");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RefreshTable()
		{
			return Factory.ExecuteBoolMethodGet(this, "RefreshTable");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedFields CalculatedFields()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedFields>(this, "CalculatedFields", NetOffice.ExcelApi.CalculatedFields.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double GetData(string name)
		{
			return Factory.ExecuteDoubleMethodGet(this, "GetData", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 ListFormulas()
		{
			return Factory.ExecuteInt32MethodGet(this, "ListFormulas");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCache PivotCache()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "PivotCache", NetOffice.ExcelApi.PivotCache.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
		/// <param name="readData">optional object readData</param>
		/// <param name="connection">optional object connection</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard()
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType, sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType, sourceData, tableDestination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", sourceType, sourceData, tableDestination, tableName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
		/// <param name="readData">optional object readData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotSelect(string name, object mode)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotSelect", name, mode);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
		/// <param name="useStandardName">optional object useStandardName</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 PivotSelect(string name, object mode, object useStandardName)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotSelect", name, mode, useStandardName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PivotSelect(string name)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotSelect", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Update()
		{
			return Factory.ExecuteInt32MethodGet(this, "Update");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">NetOffice.ExcelApi.Enums.xlPivotFormatType format</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Format(NetOffice.ExcelApi.Enums.xlPivotFormatType format)
		{
			return Factory.ExecuteInt32MethodGet(this, "Format", format);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="mode">optional NetOffice.ExcelApi.Enums.XlPTSelectionMode Mode = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PivotSelect(string name, object mode)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PivotSelect", name, mode);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PivotSelect(string name)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PivotSelect", name);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		/// <param name="field12">optional object field12</param>
		/// <param name="item12">optional object item12</param>
		/// <param name="field13">optional object field13</param>
		/// <param name="item13">optional object item13</param>
		/// <param name="field14">optional object field14</param>
		/// <param name="item14">optional object item14</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13, object field14, object item14)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13, item13, field14, item14 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, dataField);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, dataField, field1);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, dataField, field1, item1);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, dataField, field1, item1, field2);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		/// <param name="field12">optional object field12</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		/// <param name="field12">optional object field12</param>
		/// <param name="item12">optional object item12</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		/// <param name="field12">optional object field12</param>
		/// <param name="item12">optional object item12</param>
		/// <param name="field13">optional object field13</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		/// <param name="field12">optional object field12</param>
		/// <param name="item12">optional object item12</param>
		/// <param name="field13">optional object field13</param>
		/// <param name="item13">optional object item13</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13, item13 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="field1">optional object field1</param>
		/// <param name="item1">optional object item1</param>
		/// <param name="field2">optional object field2</param>
		/// <param name="item2">optional object item2</param>
		/// <param name="field3">optional object field3</param>
		/// <param name="item3">optional object item3</param>
		/// <param name="field4">optional object field4</param>
		/// <param name="item4">optional object item4</param>
		/// <param name="field5">optional object field5</param>
		/// <param name="item5">optional object item5</param>
		/// <param name="field6">optional object field6</param>
		/// <param name="item6">optional object item6</param>
		/// <param name="field7">optional object field7</param>
		/// <param name="item7">optional object item7</param>
		/// <param name="field8">optional object field8</param>
		/// <param name="item8">optional object item8</param>
		/// <param name="field9">optional object field9</param>
		/// <param name="item9">optional object item9</param>
		/// <param name="field10">optional object field10</param>
		/// <param name="item10">optional object item10</param>
		/// <param name="field11">optional object field11</param>
		/// <param name="item11">optional object item11</param>
		/// <param name="field12">optional object field12</param>
		/// <param name="item12">optional object item12</param>
		/// <param name="field13">optional object field13</param>
		/// <param name="item13">optional object item13</param>
		/// <param name="field14">optional object field14</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range GetPivotData(object dataField, object field1, object item1, object field2, object item2, object field3, object item3, object field4, object item4, object field5, object item5, object field6, object item6, object field7, object item7, object field8, object item8, object field9, object item9, object field10, object item10, object field11, object item11, object field12, object item12, object field13, object item13, object field14)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "GetPivotData", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ dataField, field1, item1, field2, item2, field3, item3, field4, item4, field5, item5, field6, item6, field7, item7, field8, item8, field9, item9, field10, item10, field11, item11, field12, item12, field13, item13, field14 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="field">object field</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="function">optional object function</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField AddDataField(object field, object caption, object function)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotField>(this, "AddDataField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType, field, caption, function);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="field">object field</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField AddDataField(object field)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotField>(this, "AddDataField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType, field);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="field">object field</param>
		/// <param name="caption">optional object caption</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField AddDataField(object field, object caption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotField>(this, "AddDataField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType, field, caption);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		/// <param name="arg30">optional object arg30</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", arg1);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", arg1, arg2, arg3);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", arg1, arg2, arg3, arg4);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Dummy15(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy15", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="measures">optional object measures</param>
		/// <param name="levels">optional object levels</param>
		/// <param name="members">optional object members</param>
		/// <param name="properties">optional object properties</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string CreateCubeFile(string file, object measures, object levels, object members, object properties)
		{
			return Factory.ExecuteStringMethodGet(this, "CreateCubeFile", new object[]{ file, measures, levels, members, properties });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string CreateCubeFile(string file)
		{
			return Factory.ExecuteStringMethodGet(this, "CreateCubeFile", file);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="measures">optional object measures</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string CreateCubeFile(string file, object measures)
		{
			return Factory.ExecuteStringMethodGet(this, "CreateCubeFile", file, measures);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="measures">optional object measures</param>
		/// <param name="levels">optional object levels</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string CreateCubeFile(string file, object measures, object levels)
		{
			return Factory.ExecuteStringMethodGet(this, "CreateCubeFile", file, measures, levels);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="measures">optional object measures</param>
		/// <param name="levels">optional object levels</param>
		/// <param name="members">optional object members</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string CreateCubeFile(string file, object measures, object levels, object members)
		{
			return Factory.ExecuteStringMethodGet(this, "CreateCubeFile", file, measures, levels, members);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ClearTable()
		{
			return Factory.ExecuteInt32MethodGet(this, "ClearTable");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="rowLayout">NetOffice.ExcelApi.Enums.XlLayoutRowType rowLayout</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 RowAxisLayout(NetOffice.ExcelApi.Enums.XlLayoutRowType rowLayout)
		{
			return Factory.ExecuteInt32MethodGet(this, "RowAxisLayout", rowLayout);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="location">NetOffice.ExcelApi.Enums.xLSubtototalLocationType location</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 SubtotalLocation(NetOffice.ExcelApi.Enums.xLSubtototalLocationType location)
		{
			return Factory.ExecuteInt32MethodGet(this, "SubtotalLocation", location);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ClearAllFilters()
		{
			return Factory.ExecuteInt32MethodGet(this, "ClearAllFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="convertFilters">bool convertFilters</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ConvertToFormulas(bool convertFilters)
		{
			return Factory.ExecuteInt32MethodGet(this, "ConvertToFormulas", convertFilters);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="conn">NetOffice.ExcelApi.WorkbookConnection conn</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ChangeConnection(NetOffice.ExcelApi.WorkbookConnection conn)
		{
			return Factory.ExecuteInt32MethodGet(this, "ChangeConnection", conn);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="pivotCache">object pivotCache</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ChangePivotCache(object pivotCache)
		{
			return Factory.ExecuteInt32MethodGet(this, "ChangePivotCache", pivotCache);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 AllocateChanges()
		{
			return Factory.ExecuteInt32MethodGet(this, "AllocateChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 CommitChanges()
		{
			return Factory.ExecuteInt32MethodGet(this, "CommitChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 DiscardChanges()
		{
			return Factory.ExecuteInt32MethodGet(this, "DiscardChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 RefreshDataSourceValues()
		{
			return Factory.ExecuteInt32MethodGet(this, "RefreshDataSourceValues");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="repeat">NetOffice.ExcelApi.Enums.XlPivotFieldRepeatLabels repeat</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 RepeatAllLabels(NetOffice.ExcelApi.Enums.XlPivotFieldRepeatLabels repeat)
		{
			return Factory.ExecuteInt32MethodGet(this, "RepeatAllLabels", repeat);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="rowline">optional object rowline</param>
		/// <param name="columnline">optional object columnline</param>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotValueCell PivotValueCell(object rowline, object columnline)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotValueCell>(this, "PivotValueCell", NetOffice.ExcelApi.PivotValueCell.LateBindingApiWrapperType, rowline, columnline);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotValueCell PivotValueCell()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotValueCell>(this, "PivotValueCell", NetOffice.ExcelApi.PivotValueCell.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="rowline">optional object rowline</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotValueCell PivotValueCell(object rowline)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotValueCell>(this, "PivotValueCell", NetOffice.ExcelApi.PivotValueCell.LateBindingApiWrapperType, rowline);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		/// <param name="pivotLine">optional object pivotLine</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillDown(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillDown", pivotItem, pivotLine);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillDown(NetOffice.ExcelApi.PivotItem pivotItem)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillDown", pivotItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		/// <param name="pivotLine">optional object pivotLine</param>
		/// <param name="levelUniqueName">optional object levelUniqueName</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillUp(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine, object levelUniqueName)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillUp", pivotItem, pivotLine, levelUniqueName);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillUp(NetOffice.ExcelApi.PivotItem pivotItem)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillUp", pivotItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		/// <param name="pivotLine">optional object pivotLine</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillUp(NetOffice.ExcelApi.PivotItem pivotItem, object pivotLine)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillUp", pivotItem, pivotLine);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		/// <param name="cubeField">NetOffice.ExcelApi.CubeField cubeField</param>
		/// <param name="pivotLine">optional object pivotLine</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillTo(NetOffice.ExcelApi.PivotItem pivotItem, NetOffice.ExcelApi.CubeField cubeField, object pivotLine)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillTo", pivotItem, cubeField, pivotLine);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="pivotItem">NetOffice.ExcelApi.PivotItem pivotItem</param>
		/// <param name="cubeField">NetOffice.ExcelApi.CubeField cubeField</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public Int32 DrillTo(NetOffice.ExcelApi.PivotItem pivotItem, NetOffice.ExcelApi.CubeField cubeField)
		{
			return Factory.ExecuteInt32MethodGet(this, "DrillTo", pivotItem, cubeField);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 15, 16)]
		public object Dummy2(object arg1, object arg2, object arg3, object arg4)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy2", arg1, arg2, arg3, arg4);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Dummy2(object arg1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy2", arg1);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Dummy2(object arg1, object arg2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy2", arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="arg1">object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Dummy2(object arg1, object arg2, object arg3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Dummy2", arg1, arg2, arg3);
		}

		#endregion

		#pragma warning restore
	}
}
