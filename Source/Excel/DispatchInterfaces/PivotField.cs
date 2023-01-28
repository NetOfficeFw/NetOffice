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
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ChildItems"/>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ChildItems"/> </remarks>
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
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.HiddenItems"/>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.HiddenItems"/> </remarks>
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
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ParentItems"/>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ParentItems"/> </remarks>
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
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Subtotals"/>
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
			Factory.ExecutePropertySet(this, "Subtotals", value, index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Subtotals
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Subtotals"/> </remarks>
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
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.VisibleItems"/>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.VisibleItems"/> </remarks>
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
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Calculation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ChildField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ChildItems"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.CurrentPage"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DataRange"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DataType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Function"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.GroupLevel"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.HiddenItems"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.LabelRange"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Name"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.NumberFormat"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Orientation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ShowAllItems"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ParentField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ParentItems"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Position"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.SourceName"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Subtotals"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.BaseField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.BaseItem"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.TotalLevels"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Value"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.VisibleItems"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DragToColumn"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DragToHide"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DragToPage"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DragToRow"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DragToData"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Formula"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.IsCalculated"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.MemoryUsed"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ServerBased"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSortOrder"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSortField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoShowType"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoShowRange"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoShowCount"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoShowField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.LayoutBlankLine"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.LayoutSubtotalLocation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.LayoutPageBreak"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.LayoutForm"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.SubtotalName"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Caption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DrilledDown"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.CubeField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.CurrentPageName"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.StandardFormula"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.HiddenItemsList"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DatabaseSort"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.IsMemberProperty"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.PropertyParentField"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.PropertyOrder"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.EnableItemSelection"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.CurrentPageList"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Hidden"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.UseMemberPropertyAsCaption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.MemberPropertyCaption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DisplayAsTooltip"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DisplayInReport"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DisplayAsCaption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.LayoutCompactRow"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.IncludeNewItemsInFilter"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.VisibleItemsList"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.PivotFilters"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSortPivotLine"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSortCustomSubtotal"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ShowingInAxis"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.EnableMultiplePageItems"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AllItemsVisible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.SourceCaption"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ShowDetail"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.RepeatLabels"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.PivotItems"/> </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotItems(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.PivotItems"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotItems()
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotItems");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.CalculatedItems"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CalculatedItems CalculatedItems()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedItems>(this, "CalculatedItems", NetOffice.ExcelApi.CalculatedItems.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSort"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSort"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoSort"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AutoShow"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AddPageItem"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.AddPageItem"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.DrillTo"/> </remarks>
		/// <param name="field">string field</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void DrillTo(string field)
		{
			 Factory.ExecuteMethod(this, "DrillTo", field);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ClearManualFilter"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearManualFilter()
		{
			 Factory.ExecuteMethod(this, "ClearManualFilter");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ClearAllFilters"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearAllFilters()
		{
			 Factory.ExecuteMethod(this, "ClearAllFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ClearValueFilters"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearValueFilters()
		{
			 Factory.ExecuteMethod(this, "ClearValueFilters");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotField.ClearLabelFilters"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearLabelFilters()
		{
			 Factory.ExecuteMethod(this, "ClearLabelFilters");
		}

		#endregion

		#pragma warning restore
	}
}
