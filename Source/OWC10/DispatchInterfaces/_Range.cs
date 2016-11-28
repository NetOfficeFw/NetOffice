using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// _Range
	///</summary>
	public class _Range_ : COMObject
	{
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Address(object rowAbsolute)
		{
			return get_Address(rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Address(object rowAbsolute, object columnAbsolute)
		{
			return get_Address(rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle ReferenceStyle</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle ReferenceStyle</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle, external);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.OWC10Api.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		/// <param name="columnOffset">optional object ColumnOffset</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Offset(object rowOffset, object columnOffset)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowOffset, columnOffset);
			object returnItem = Invoker.PropertyGet(this, "Offset", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Offset
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		/// <param name="columnOffset">optional object ColumnOffset</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Offset(object rowOffset, object columnOffset)
		{
			return get_Offset(rowOffset, columnOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Offset(object rowOffset)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowOffset);
			object returnItem = Invoker.PropertyGet(this, "Offset", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Offset
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Offset(object rowOffset)
		{
			return get_Offset(rowOffset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="rangeValueDataType">optional object RangeValueDataType</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Value(object rangeValueDataType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rangeValueDataType);
			object returnItem = Invoker.PropertyGet(this, "Value", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object RangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Value(object rangeValueDataType, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rangeValueDataType);
			Invoker.PropertySet(this, "Value", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Value
		/// </summary>
		/// <param name="rangeValueDataType">optional object RangeValueDataType</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Value(object rangeValueDataType)
		{
			return get_Value(rangeValueDataType);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface _Range 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Range : _Range_ ,IEnumerable<object>
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Range(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="row">optional object Row</param>
		/// <param name="column">optional object Column</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object row, object column]
		{
			get
			{
			object[] paramsArray = Invoker.ValidateParamsArray(row, column);
			object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
			}
			set
			{
			object[] paramsArray = Invoker.ValidateParamsArray(row, column);
			Invoker.PropertySet(this, "_Default", paramsArray, value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="row">optional object Row</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object row]
		{
			get
			{
			object[] paramsArray = Invoker.ValidateParamsArray(row);
			object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
			}
			set
			{
			object[] paramsArray = Invoker.ValidateParamsArray(row);
			Invoker.PropertySet(this, "_Default", paramsArray, value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Address
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OWC10Api.ISpreadsheet newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api.ISpreadsheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Borders Borders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Borders", paramsArray);
				NetOffice.OWC10Api.Borders newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Borders.LateBindingApiWrapperType) as NetOffice.OWC10Api.Borders;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Column
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Column", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object ColumnWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnWidth", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ColumnWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range CurrentArray
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentArray", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range CurrentRegion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentRegion", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection Direction</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_End(NetOffice.OWC10Api.Enums.XlDirection direction)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			object returnItem = Invoker.PropertyGet(this, "End", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_End
		/// </summary>
		/// <param name="direction">NetOffice.OWC10Api.Enums.XlDirection Direction</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range End(NetOffice.OWC10Api.Enums.XlDirection direction)
		{
			return get_End(direction);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range EntireColumn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EntireColumn", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range EntireRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EntireRow", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				NetOffice.OWC10Api.Font newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Font.LateBindingApiWrapperType) as NetOffice.OWC10Api.Font;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Formula
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Formula", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Formula", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object FormulaArray
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaArray", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaArray", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object FormulaLocal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaLocal", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaLocal", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object HasArray
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasArray", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object HasFormula
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasFormula", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool Hidden
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hidden", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Hidden", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object HorizontalAlignment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HorizontalAlignment", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HorizontalAlignment", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string HTMLData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLData", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Hyperlink Hyperlink
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlink", paramsArray);
				NetOffice.OWC10Api.Hyperlink newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Hyperlink.LateBindingApiWrapperType) as NetOffice.OWC10Api.Hyperlink;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Interior Interior
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Interior", paramsArray);
				NetOffice.OWC10Api.Interior newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Interior.LateBindingApiWrapperType) as NetOffice.OWC10Api.Interior;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Left
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Left", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Locked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Locked", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Locked", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range MergeArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MergeArea", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object MergeCells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MergeCells", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MergeCells", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Name Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				NetOffice.OWC10Api.Name newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Name.LateBindingApiWrapperType) as NetOffice.OWC10Api.Name;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Next
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Next", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object NumberFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NumberFormat", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NumberFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Offset
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Offset", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.OWC10Api.Worksheet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType) as NetOffice.OWC10Api.Worksheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object PrefixCharacter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrefixCharacter", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Previous
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Previous", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1, object cell2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1, cell2);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object ReadingOrder
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadingOrder", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadingOrder", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Row
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Row", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object RowHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowHeight", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RowHeight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Text", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Top
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Top", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object UseStandardHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseStandardHeight", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseStandardHeight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object UseStandardWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseStandardWidth", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseStandardWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Value
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Value", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Value", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Value2
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Value2", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Value2", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object VerticalAlignment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VerticalAlignment", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "VerticalAlignment", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Worksheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Worksheet", paramsArray);
				NetOffice.OWC10Api.Worksheet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType) as NetOffice.OWC10Api.Worksheet;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		/// <param name="_operator">optional object Operator</param>
		/// <param name="criteria2">optional object Criteria2</param>
		/// <param name="visibleDropDown">optional object VisibleDropDown</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1, _operator, criteria2, visibleDropDown);
			Invoker.Method(this, "AutoFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFilter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFilter(object field)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field);
			Invoker.Method(this, "AutoFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFilter(object field, object criteria1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1);
			Invoker.Method(this, "AutoFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		/// <param name="_operator">optional object Operator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFilter(object field, object criteria1, object _operator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1, _operator);
			Invoker.Method(this, "AutoFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		/// <param name="_operator">optional object Operator</param>
		/// <param name="criteria2">optional object Criteria2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFilter(object field, object criteria1, object _operator, object criteria2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1, _operator, criteria2);
			Invoker.Method(this, "AutoFilter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoFit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoFit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object Color</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void BorderAround(object lineStyle, object weight, object colorIndex, object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex, color);
			Invoker.Method(this, "BorderAround", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void BorderAround()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BorderAround", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void BorderAround(object lineStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle);
			Invoker.Method(this, "BorderAround", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void BorderAround(object lineStyle, object weight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight);
			Invoker.Method(this, "BorderAround", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.OWC10Api.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.OWC10Api.Enums.XlColorIndex ColorIndex = -4105</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void BorderAround(object lineStyle, object weight, object colorIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex);
			Invoker.Method(this, "BorderAround", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Calculate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Calculate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Clear()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Clear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void ClearFormats()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearFormats", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void ClearContents()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearContents", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Copy(object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="data">object Data</param>
		/// <param name="maxRows">optional object MaxRows</param>
		/// <param name="maxColumns">optional object MaxColumns</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 CopyFromRecordset(object data, object maxRows, object maxColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data, maxRows, maxColumns);
			object returnItem = Invoker.MethodReturn(this, "CopyFromRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="data">object Data</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 CopyFromRecordset(object data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data);
			object returnItem = Invoker.MethodReturn(this, "CopyFromRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="data">object Data</param>
		/// <param name="maxRows">optional object MaxRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 CopyFromRecordset(object data, object maxRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data, maxRows);
			object returnItem = Invoker.MethodReturn(this, "CopyFromRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Cut(object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="shift">optional object Shift</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Delete(object shift)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shift);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void FillDown()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FillDown", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void FillRight()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FillRight", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="matchByte">optional object MatchByte</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.OWC10Api.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range FindNext(object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after);
			object returnItem = Invoker.MethodReturn(this, "FindNext", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range FindNext()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindNext", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range FindPrevious(object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after);
			object returnItem = Invoker.MethodReturn(this, "FindPrevious", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range FindPrevious()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindPrevious", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="shift">optional object Shift</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Insert(object shift)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shift);
			Invoker.Method(this, "Insert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Insert()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Insert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="file">string File</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		/// <param name="textQualifier">optional string TextQualifier = \042</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void LoadText(string file, object delimiters, object consecutiveDelimAsOne, object textQualifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(file, delimiters, consecutiveDelimAsOne, textQualifier);
			Invoker.Method(this, "LoadText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="file">string File</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void LoadText(string file)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(file);
			Invoker.Method(this, "LoadText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="file">string File</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void LoadText(string file, object delimiters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(file, delimiters);
			Invoker.Method(this, "LoadText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="file">string File</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void LoadText(string file, object delimiters, object consecutiveDelimAsOne)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(file, delimiters, consecutiveDelimAsOne);
			Invoker.Method(this, "LoadText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="across">optional object Across</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Merge(object across)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(across);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Merge()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		/// <param name="textQualifier">optional string TextQualifier = \042</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void ParseText(string text, object delimiters, object consecutiveDelimAsOne, object textQualifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, delimiters, consecutiveDelimAsOne, textQualifier);
			Invoker.Method(this, "ParseText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ParseText(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "ParseText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ParseText(string text, object delimiters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, delimiters);
			Invoker.Method(this, "ParseText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="delimiters">optional string Delimiters = </param>
		/// <param name="consecutiveDelimAsOne">optional bool ConsecutiveDelimAsOne = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ParseText(string text, object delimiters, object consecutiveDelimAsOne)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, delimiters, consecutiveDelimAsOne);
			Invoker.Method(this, "ParseText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Show()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Show", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="columnKey">optional Int32 ColumnKey = 1</param>
		/// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
		/// <param name="header">optional NetOffice.OWC10Api.Enums.XlYesNoGuess Header = 2</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Sort(object columnKey, object order, object header)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(columnKey, order, header);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Sort()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="columnKey">optional Int32 ColumnKey = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Sort(object columnKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(columnKey);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="columnKey">optional Int32 ColumnKey = 1</param>
		/// <param name="order">optional NetOffice.OWC10Api.Enums.XlSortOrder Order = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Sort(object columnKey, object order)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(columnKey, order);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void UnMerge()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UnMerge", paramsArray);
		}

		#endregion

       #region IEnumerable<object> Member
        
        /// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
       public IEnumerator<object> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (object item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute OWC10, 1
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}