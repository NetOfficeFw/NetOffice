using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// PivotField
	/// </summary>
	[SyntaxBypass]
 	public class PivotField_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotField_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotField_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838618.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ChildItems(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "ChildItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ChildItems
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838618.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_ChildItems")]
		public object ChildItems(object index)
		{
			return get_ChildItems(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840567.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HiddenItems(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "HiddenItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_HiddenItems
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840567.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_HiddenItems")]
		public object HiddenItems(object index)
		{
			return get_HiddenItems(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196532.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ParentItems(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "ParentItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ParentItems
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196532.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_ParentItems")]
		public object ParentItems(object index)
		{
			return get_ParentItems(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193947.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Subtotals(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Subtotals", index);
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Subtotals(object index, object value)
		{
			Factory.ExecutePropertySet(this, "Subtotals", index, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Subtotals
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193947.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Subtotals")]
		public object Subtotals(object index)
		{
			return get_Subtotals(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841013.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_VisibleItems(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "VisibleItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_VisibleItems
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841013.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_VisibleItems")]
		public object VisibleItems(object index)
		{
			return get_VisibleItems(index);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface PivotField 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820762.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotField : PivotField_
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
                    _type = typeof(PivotField);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotField(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotField(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197845.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198166.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834483.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837646.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPivotFieldCalculation Calculation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotFieldCalculation>(this, "Calculation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Calculation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836464.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField ChildField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "ChildField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838618.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ChildItems
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ChildItems");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197857.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CurrentPage
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CurrentPage");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CurrentPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194280.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DataRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836207.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPivotFieldDataType DataType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotFieldDataType>(this, "DataType");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835023.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlConsolidationFunction Function
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlConsolidationFunction>(this, "Function");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Function", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841193.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object GroupLevel
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "GroupLevel");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840567.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object HiddenItems
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HiddenItems");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838465.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range LabelRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "LabelRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193604.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195116.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string NumberFormat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NumberFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194695.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPivotFieldOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotFieldOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835616.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197900.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField ParentField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "ParentField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196532.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ParentItems
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ParentItems");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835912.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Position
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Position");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839397.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string SourceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193947.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Subtotals
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Subtotals");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Subtotals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822779.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BaseField
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaseField");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaseField", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194011.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BaseItem
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaseItem");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaseItem", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841161.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TotalLevels
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TotalLevels");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838634.aspx </remarks>
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841013.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object VisibleItems
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "VisibleItems");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194893.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195313.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToHide
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToHide");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToHide", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196930.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToPage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840765.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196219.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821210.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Formula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837059.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool IsCalculated
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCalculated");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835860.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 MemoryUsed
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MemoryUsed");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835590.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ServerBased
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ServerBased");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerBased", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838026.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 AutoSortOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoSortOrder");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840859.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string AutoSortField
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AutoSortField");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837571.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 AutoShowType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoShowType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838186.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 AutoShowRange
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoShowRange");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838416.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 AutoShowCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoShowCount");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835286.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string AutoShowField
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AutoShowField");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835259.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool LayoutBlankLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutBlankLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutBlankLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834311.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xLSubtototalLocationType LayoutSubtotalLocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xLSubtototalLocationType>(this, "LayoutSubtotalLocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutSubtotalLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196893.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool LayoutPageBreak
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutPageBreak");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutPageBreak", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821507.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlLayoutFormType LayoutForm
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlLayoutFormType>(this, "LayoutForm");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839840.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string SubtotalName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SubtotalName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubtotalName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834625.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822783.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DrilledDown
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DrilledDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrilledDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839541.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CubeField CubeField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CubeField>(this, "CubeField", NetOffice.ExcelApi.CubeField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839223.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string CurrentPageName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentPageName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentPageName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194277.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string StandardFormula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StandardFormula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StandardFormula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195478.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object HiddenItemsList
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HiddenItemsList");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HiddenItemsList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194504.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool DatabaseSort
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DatabaseSort");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatabaseSort", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840159.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool IsMemberProperty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsMemberProperty");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836482.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField PropertyParentField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "PropertyParentField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838259.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 PropertyOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PropertyOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PropertyOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837828.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EnableItemSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableItemSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableItemSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197170.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object CurrentPageList
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CurrentPageList");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CurrentPageList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838817.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool Hidden
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Hidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198159.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool UseMemberPropertyAsCaption
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseMemberPropertyAsCaption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseMemberPropertyAsCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822155.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string MemberPropertyCaption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MemberPropertyCaption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MemberPropertyCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197797.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DisplayAsTooltip
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAsTooltip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAsTooltip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821563.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DisplayInReport
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayInReport");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayInReport", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838030.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DisplayAsCaption
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAsCaption");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840966.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool LayoutCompactRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutCompactRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutCompactRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837852.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool IncludeNewItemsInFilter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeNewItemsInFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeNewItemsInFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839948.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object VisibleItemsList
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "VisibleItemsList");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "VisibleItemsList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841070.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilters PivotFilters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFilters>(this, "PivotFilters", NetOffice.ExcelApi.PivotFilters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194197.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotLine AutoSortPivotLine
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotLine>(this, "AutoSortPivotLine", NetOffice.ExcelApi.PivotLine.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837075.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 AutoSortCustomSubtotal
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoSortCustomSubtotal");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821193.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowingInAxis
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowingInAxis");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836478.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool EnableMultiplePageItems
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableMultiplePageItems");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableMultiplePageItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835827.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool AllItemsVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllItemsVisible");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837384.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string SourceCaption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceCaption");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196520.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowDetail
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDetail");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDetail", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837644.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool RepeatLabels
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RepeatLabels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RepeatLabels", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821539.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotItems(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821539.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotItems()
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotItems");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835297.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedItems CalculatedItems()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedItems>(this, "CalculatedItems", NetOffice.ExcelApi.CalculatedItems.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197782.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834371.aspx </remarks>
		/// <param name="order">Int32 order</param>
		/// <param name="field">string field</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AutoSort(Int32 order, string field)
		{
			 Factory.ExecuteMethod(this, "AutoSort", order, field);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834371.aspx </remarks>
		/// <param name="order">Int32 order</param>
		/// <param name="field">string field</param>
		/// <param name="pivotLine">optional object pivotLine</param>
		/// <param name="customSubtotal">optional object customSubtotal</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void AutoSort(Int32 order, string field, object pivotLine, object customSubtotal)
		{
			 Factory.ExecuteMethod(this, "AutoSort", order, field, pivotLine, customSubtotal);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834371.aspx </remarks>
		/// <param name="order">Int32 order</param>
		/// <param name="field">string field</param>
		/// <param name="pivotLine">optional object pivotLine</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void AutoSort(Int32 order, string field, object pivotLine)
		{
			 Factory.ExecuteMethod(this, "AutoSort", order, field, pivotLine);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835595.aspx </remarks>
		/// <param name="type">Int32 type</param>
		/// <param name="range">Int32 range</param>
		/// <param name="count">Int32 count</param>
		/// <param name="field">string field</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AutoShow(Int32 type, Int32 range, Int32 count, string field)
		{
			 Factory.ExecuteMethod(this, "AutoShow", type, range, count, field);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838994.aspx </remarks>
		/// <param name="item">string item</param>
		/// <param name="clearList">optional object clearList</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void AddPageItem(string item, object clearList)
		{
			 Factory.ExecuteMethod(this, "AddPageItem", item, clearList);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838994.aspx </remarks>
		/// <param name="item">string item</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void AddPageItem(string item)
		{
			 Factory.ExecuteMethod(this, "AddPageItem", item);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="order">Int32 order</param>
		/// <param name="field">string field</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _AutoSort(Int32 order, string field)
		{
			 Factory.ExecuteMethod(this, "_AutoSort", order, field);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837088.aspx </remarks>
		/// <param name="field">string field</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void DrillTo(string field)
		{
			 Factory.ExecuteMethod(this, "DrillTo", field);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822577.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearManualFilter()
		{
			 Factory.ExecuteMethod(this, "ClearManualFilter");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821564.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearAllFilters()
		{
			 Factory.ExecuteMethod(this, "ClearAllFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835314.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearValueFilters()
		{
			 Factory.ExecuteMethod(this, "ClearValueFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197790.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearLabelFilters()
		{
			 Factory.ExecuteMethod(this, "ClearLabelFilters");
		}

		#endregion

		#pragma warning restore
	}
}
