using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// _Range
	/// </summary>
	[SyntaxBypass]
 	public class _Range_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Range_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Range_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// <param name="relativeTo">optional object relativeTo</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", new object[]{ rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// <param name="relativeTo">optional object relativeTo</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Address")]
		public string Address(object rowAbsolute)
		{
			return get_Address(rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute)
		{
			return get_Address(rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object rowOffset</param>
		/// <param name="columnOffset">optional object columnOffset</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Offset(object rowOffset, object columnOffset)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Offset", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, rowOffset, columnOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Offset
		/// </summary>
		/// <param name="rowOffset">optional object rowOffset</param>
		/// <param name="columnOffset">optional object columnOffset</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Offset")]
		public NetOffice.OWC10Api._Range Offset(object rowOffset, object columnOffset)
		{
			return get_Offset(rowOffset, columnOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object rowOffset</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Offset(object rowOffset)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Offset", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, rowOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Offset
		/// </summary>
		/// <param name="rowOffset">optional object rowOffset</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Offset")]
		public NetOffice.OWC10Api._Range Offset(object rowOffset)
		{
			return get_Offset(rowOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="rangeValueDataType">optional object rangeValueDataType</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Value(object rangeValueDataType)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Value", rangeValueDataType);
		}

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Value(object rangeValueDataType, object value)
		{
			Factory.ExecutePropertySet(this, "Value", rangeValueDataType, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Value
		/// </summary>
		/// <param name="rangeValueDataType">optional object rangeValueDataType</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Value")]
		public object Value(object rangeValueDataType)
		{
			return get_Value(rangeValueDataType);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface _Range 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class _Range : _Range_, IEnumerableProvider<object>
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
                    _type = typeof(_Range);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Range(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Range(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range(string progId) : base(progId)
		{
		}

        #endregion

        #region Properties
        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// Custom Indexer
        /// </summary>
        /// <param name="row">optional object row</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public object this[object row]
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "_Default", row);
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "_Default", value, row);
			}
		}

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="row">optional object row</param>
        /// <param name="column">optional object column</param>
        [SupportByVersion("OWC10", 1)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public object this[object row, object column]
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "_Default", row, column);
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "_Default", value, row, column);
			}
		}

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
		public string Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Address");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Borders Borders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Borders>(this, "Borders", NetOffice.OWC10Api.Borders.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Cells
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Cells");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Column
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Column");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Columns
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Columns");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object ColumnWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ColumnWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ColumnWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range CurrentArray
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "CurrentArray");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range CurrentRegion
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "CurrentRegion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection direction</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_End(NetOffice.OWC10Api.Enums.XlDirection direction)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "End", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, direction);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_End
		/// </summary>
		/// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection direction</param>
		[SupportByVersion("OWC10", 1), Redirect("get_End")]
		public NetOffice.OWC10Api._Range End(NetOffice.OWC10Api.Enums.XlDirection direction)
		{
			return get_End(direction);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range EntireColumn
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "EntireColumn");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range EntireRow
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "EntireRow");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Font Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Font>(this, "Font", NetOffice.OWC10Api.Font.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Formula
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Formula");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object FormulaArray
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FormulaArray");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FormulaArray", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object FormulaLocal
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FormulaLocal");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FormulaLocal", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object HasArray
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HasArray");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object HasFormula
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HasFormula");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Height
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object HorizontalAlignment
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HorizontalAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string HTMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HTMLData");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Hyperlink Hyperlink
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Hyperlink>(this, "Hyperlink", NetOffice.OWC10Api.Hyperlink.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Interior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Interior>(this, "Interior", NetOffice.OWC10Api.Interior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Left
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Locked
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Locked");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Locked", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range MergeArea
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "MergeArea");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object MergeCells
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MergeCells");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MergeCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Name Name
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Name>(this, "Name", NetOffice.OWC10Api.Name.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Next
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Next");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object NumberFormat
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "NumberFormat");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "NumberFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Offset
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Offset");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Parent", NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object PrefixCharacter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrefixCharacter");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Previous
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Previous");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1, object cell2)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public NetOffice.OWC10Api._Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public NetOffice.OWC10Api._Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object ReadingOrder
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ReadingOrder");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ReadingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Row
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Row");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object RowHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RowHeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Rows
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Rows");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Text
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Text");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Top
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object UseStandardHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "UseStandardHeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "UseStandardHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object UseStandardWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "UseStandardWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "UseStandardWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Value
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Value2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value2", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object VerticalAlignment
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "VerticalAlignment");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "VerticalAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object Width
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Worksheet
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Worksheet", NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="criteria2">optional object criteria2</param>
		/// <param name="visibleDropDown">optional object visibleDropDown</param>
		[SupportByVersion("OWC10", 1)]
		public void AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown)
		{
			 Factory.ExecuteMethod(this, "AutoFilter", new object[]{ field, criteria1, _operator, criteria2, visibleDropDown });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AutoFilter()
		{
			 Factory.ExecuteMethod(this, "AutoFilter");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">optional object field</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AutoFilter(object field)
		{
			 Factory.ExecuteMethod(this, "AutoFilter", field);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AutoFilter(object field, object criteria1)
		{
			 Factory.ExecuteMethod(this, "AutoFilter", field, criteria1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		/// <param name="_operator">optional object operator</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AutoFilter(object field, object criteria1, object _operator)
		{
			 Factory.ExecuteMethod(this, "AutoFilter", field, criteria1, _operator);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="criteria2">optional object criteria2</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AutoFilter(object field, object criteria1, object _operator, object criteria2)
		{
			 Factory.ExecuteMethod(this, "AutoFilter", field, criteria1, _operator, criteria2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void AutoFit()
		{
			 Factory.ExecuteMethod(this, "AutoFit");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object color</param>
		[SupportByVersion("OWC10", 1)]
		public void BorderAround(object lineStyle, object weight, object colorIndex, object color)
		{
			 Factory.ExecuteMethod(this, "BorderAround", lineStyle, weight, colorIndex, color);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void BorderAround()
		{
			 Factory.ExecuteMethod(this, "BorderAround");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void BorderAround(object lineStyle)
		{
			 Factory.ExecuteMethod(this, "BorderAround", lineStyle);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void BorderAround(object lineStyle, object weight)
		{
			 Factory.ExecuteMethod(this, "BorderAround", lineStyle, weight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void BorderAround(object lineStyle, object weight, object colorIndex)
		{
			 Factory.ExecuteMethod(this, "BorderAround", lineStyle, weight, colorIndex);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Calculate()
		{
			 Factory.ExecuteMethod(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ClearFormats()
		{
			 Factory.ExecuteMethod(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ClearContents()
		{
			 Factory.ExecuteMethod(this, "ClearContents");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("OWC10", 1)]
		public void Copy(object destination)
		{
			 Factory.ExecuteMethod(this, "Copy", destination);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="data">object data</param>
		/// <param name="maxRows">optional object maxRows</param>
		/// <param name="maxColumns">optional object maxColumns</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 CopyFromRecordset(object data, object maxRows, object maxColumns)
		{
			return Factory.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows, maxColumns);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="data">object data</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public Int32 CopyFromRecordset(object data)
		{
			return Factory.ExecuteInt32MethodGet(this, "CopyFromRecordset", data);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="data">object data</param>
		/// <param name="maxRows">optional object maxRows</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public Int32 CopyFromRecordset(object data, object maxRows)
		{
			return Factory.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("OWC10", 1)]
		public void Cut(object destination)
		{
			 Factory.ExecuteMethod(this, "Cut", destination);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="shift">optional object shift</param>
		[SupportByVersion("OWC10", 1)]
		public void Delete(object shift)
		{
			 Factory.ExecuteMethod(this, "Delete", shift);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void FillDown()
		{
			 Factory.ExecuteMethod(this, "FillDown");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void FillRight()
		{
			 Factory.ExecuteMethod(this, "FillRight");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchByte">optional object matchByte</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what, after);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what, after, lookIn);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", what, after, lookIn, lookAt);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[]{ what, after, lookIn, lookAt, searchOrder });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "Find", new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="after">optional object after</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range FindNext(object after)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindNext", after);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range FindNext()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindNext");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="after">optional object after</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range FindPrevious(object after)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindPrevious", after);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api._Range FindPrevious()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "FindPrevious");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="shift">optional object shift</param>
		[SupportByVersion("OWC10", 1)]
		public void Insert(object shift)
		{
			 Factory.ExecuteMethod(this, "Insert", shift);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Insert()
		{
			 Factory.ExecuteMethod(this, "Insert");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		/// <param name="textQualifier">optional string TextQualifier = \042</param>
		[SupportByVersion("OWC10", 1)]
		public void LoadText(string file, object delimiters, object consecutiveDelimAsOne, object textQualifier)
		{
			 Factory.ExecuteMethod(this, "LoadText", file, delimiters, consecutiveDelimAsOne, textQualifier);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="file">string file</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LoadText(string file)
		{
			 Factory.ExecuteMethod(this, "LoadText", file);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LoadText(string file, object delimiters)
		{
			 Factory.ExecuteMethod(this, "LoadText", file, delimiters);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="file">string file</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LoadText(string file, object delimiters, object consecutiveDelimAsOne)
		{
			 Factory.ExecuteMethod(this, "LoadText", file, delimiters, consecutiveDelimAsOne);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="across">optional object across</param>
		[SupportByVersion("OWC10", 1)]
		public void Merge(object across)
		{
			 Factory.ExecuteMethod(this, "Merge", across);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Merge()
		{
			 Factory.ExecuteMethod(this, "Merge");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		/// <param name="textQualifier">optional string TextQualifier = \042</param>
		[SupportByVersion("OWC10", 1)]
		public void ParseText(string text, object delimiters, object consecutiveDelimAsOne, object textQualifier)
		{
			 Factory.ExecuteMethod(this, "ParseText", text, delimiters, consecutiveDelimAsOne, textQualifier);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ParseText(string text)
		{
			 Factory.ExecuteMethod(this, "ParseText", text);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ParseText(string text, object delimiters)
		{
			 Factory.ExecuteMethod(this, "ParseText", text, delimiters);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ParseText(string text, object delimiters, object consecutiveDelimAsOne)
		{
			 Factory.ExecuteMethod(this, "ParseText", text, delimiters, consecutiveDelimAsOne);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Show()
		{
			 Factory.ExecuteMethod(this, "Show");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="columnKey">optional Int32 ColumnKey = 1</param>
		/// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
		/// <param name="header">optional NetOffice.OWC10Api.Enums.XlYesNoGuess Header = 2</param>
		[SupportByVersion("OWC10", 1)]
		public void Sort(object columnKey, object order, object header)
		{
			 Factory.ExecuteMethod(this, "Sort", columnKey, order, header);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Sort()
		{
			 Factory.ExecuteMethod(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="columnKey">optional Int32 ColumnKey = 1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Sort(object columnKey)
		{
			 Factory.ExecuteMethod(this, "Sort", columnKey);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="columnKey">optional Int32 ColumnKey = 1</param>
		/// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Sort(object columnKey, object order)
		{
			 Factory.ExecuteMethod(this, "Sort", columnKey, order);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void UnMerge()
		{
			 Factory.ExecuteMethod(this, "UnMerge");
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, true);
		}

		#endregion

		#pragma warning restore
	}
}