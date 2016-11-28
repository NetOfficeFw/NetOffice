using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// IRange
	///</summary>
	public class IRange_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IRange_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Address(object rowAbsolute)
		{
			return get_Address(rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Address(object rowAbsolute, object columnAbsolute)
		{
			return get_Address(rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle, external);
			object returnItem = Invoker.PropertyGet(this, "Address", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
			object returnItem = Invoker.PropertyGet(this, "AddressLocal", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		/// <param name="relativeTo">optional object RelativeTo</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute);
			object returnItem = Invoker.PropertyGet(this, "AddressLocal", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal(object rowAbsolute)
		{
			return get_AddressLocal(rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute);
			object returnItem = Invoker.PropertyGet(this, "AddressLocal", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal(object rowAbsolute, object columnAbsolute)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle);
			object returnItem = Invoker.PropertyGet(this, "AddressLocal", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowAbsolute, columnAbsolute, referenceStyle, external);
			object returnItem = Invoker.PropertyGet(this, "AddressLocal", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <param name="rowAbsolute">optional object RowAbsolute</param>
		/// <param name="columnAbsolute">optional object ColumnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle ReferenceStyle</param>
		/// <param name="external">optional object External</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional object Start</param>
		/// <param name="length">optional object Length</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Characters get_Characters(object start, object length)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(start, length);
			object returnItem = Invoker.PropertyGet(this, "Characters", paramsArray);
			NetOffice.ExcelApi.Characters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Characters.LateBindingApiWrapperType) as NetOffice.ExcelApi.Characters;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Characters
		/// </summary>
		/// <param name="start">optional object Start</param>
		/// <param name="length">optional object Length</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Characters Characters(object start, object length)
		{
			return get_Characters(start, length);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional object Start</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Characters get_Characters(object start)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.PropertyGet(this, "Characters", paramsArray);
			NetOffice.ExcelApi.Characters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Characters.LateBindingApiWrapperType) as NetOffice.ExcelApi.Characters;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Characters
		/// </summary>
		/// <param name="start">optional object Start</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Characters Characters(object start)
		{
			return get_Characters(start);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		/// <param name="columnOffset">optional object ColumnOffset</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Offset(object rowOffset, object columnOffset)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowOffset, columnOffset);
			object returnItem = Invoker.PropertyGet(this, "Offset", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Offset
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		/// <param name="columnOffset">optional object ColumnOffset</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Offset(object rowOffset, object columnOffset)
		{
			return get_Offset(rowOffset, columnOffset);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Offset(object rowOffset)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowOffset);
			object returnItem = Invoker.PropertyGet(this, "Offset", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Offset
		/// </summary>
		/// <param name="rowOffset">optional object RowOffset</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Offset(object rowOffset)
		{
			return get_Offset(rowOffset);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowSize">optional object RowSize</param>
		/// <param name="columnSize">optional object ColumnSize</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Resize(object rowSize, object columnSize)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowSize, columnSize);
			object returnItem = Invoker.PropertyGet(this, "Resize", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Resize
		/// </summary>
		/// <param name="rowSize">optional object RowSize</param>
		/// <param name="columnSize">optional object ColumnSize</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Resize(object rowSize, object columnSize)
		{
			return get_Resize(rowSize, columnSize);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowSize">optional object RowSize</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Resize(object rowSize)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(rowSize);
			object returnItem = Invoker.PropertyGet(this, "Resize", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Resize
		/// </summary>
		/// <param name="rowSize">optional object RowSize</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Resize(object rowSize)
		{
			return get_Resize(rowSize);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="rangeValueDataType">optional object RangeValueDataType</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
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
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object RangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Value(object rangeValueDataType, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rangeValueDataType);
			Invoker.PropertySet(this, "Value", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Alias for get_Value
		/// </summary>
		/// <param name="rangeValueDataType">optional object RangeValueDataType</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Value(object rangeValueDataType)
		{
			return get_Value(rangeValueDataType);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// Interface IRange 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IRange : COMObject ,IEnumerable<object>
	{
		#pragma warning disable
		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
                    _type = typeof(IRange);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IRange(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IRange(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCreator)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AddIndent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AddIndent", paramsArray);
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
				Invoker.PropertySet(this, "AddIndent", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AddressLocal", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Areas Areas
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Areas", paramsArray);
				NetOffice.ExcelApi.Areas newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Areas.LateBindingApiWrapperType) as NetOffice.ExcelApi.Areas;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Borders Borders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Borders", paramsArray);
				NetOffice.ExcelApi.Borders newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Borders.LateBindingApiWrapperType) as NetOffice.ExcelApi.Borders;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Characters Characters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Characters", paramsArray);
				NetOffice.ExcelApi.Characters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Characters.LateBindingApiWrapperType) as NetOffice.ExcelApi.Characters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range CurrentArray
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentArray", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range CurrentRegion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentRegion", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="rowIndex">optional object RowIndex</param>
		/// <param name="columnIndex">optional object ColumnIndex</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object rowIndex, object columnIndex]
		{
			get
			{
			object[] paramsArray = Invoker.ValidateParamsArray(rowIndex, columnIndex);
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
			object[] paramsArray = Invoker.ValidateParamsArray(rowIndex, columnIndex);
			Invoker.PropertySet(this, "_Default", paramsArray, value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="rowIndex">optional object RowIndex</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object rowIndex]
		{
			get
			{
			object[] paramsArray = Invoker.ValidateParamsArray(rowIndex);
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
			object[] paramsArray = Invoker.ValidateParamsArray(rowIndex);
			Invoker.PropertySet(this, "_Default", paramsArray, value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Dependents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dependents", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DirectDependents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DirectDependents", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DirectPrecedents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DirectPrecedents", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection Direction</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_End(NetOffice.ExcelApi.Enums.XlDirection direction)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			object returnItem = Invoker.PropertyGet(this, "End", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_End
		/// </summary>
		/// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection Direction</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range End(NetOffice.ExcelApi.Enums.XlDirection direction)
		{
			return get_End(direction);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range EntireColumn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EntireColumn", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range EntireRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EntireRow", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				NetOffice.ExcelApi.Font newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Font.LateBindingApiWrapperType) as NetOffice.ExcelApi.Font;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlFormulaLabel FormulaLabel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaLabel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlFormulaLabel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaLabel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FormulaHidden
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaHidden", paramsArray);
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
				Invoker.PropertySet(this, "FormulaHidden", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FormulaR1C1
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaR1C1", paramsArray);
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
				Invoker.PropertySet(this, "FormulaR1C1", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FormulaR1C1Local
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaR1C1Local", paramsArray);
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
				Invoker.PropertySet(this, "FormulaR1C1Local", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Hidden
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hidden", paramsArray);
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
				Invoker.PropertySet(this, "Hidden", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object IndentLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IndentLevel", paramsArray);
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
				Invoker.PropertySet(this, "IndentLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Interior Interior
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Interior", paramsArray);
				NetOffice.ExcelApi.Interior newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Interior.LateBindingApiWrapperType) as NetOffice.ExcelApi.Interior;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 ListHeaderRows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListHeaderRows", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlLocationInTable LocationInTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LocationInTable", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlLocationInTable)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range MergeArea
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MergeArea", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
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
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Next
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Next", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object NumberFormatLocal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NumberFormatLocal", paramsArray);
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
				Invoker.PropertySet(this, "NumberFormatLocal", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Offset
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Offset", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Orientation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Orientation", paramsArray);
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
				Invoker.PropertySet(this, "Orientation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object OutlineLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OutlineLevel", paramsArray);
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
				Invoker.PropertySet(this, "OutlineLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 PageBreak
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageBreak", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PageBreak", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField PivotField
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotField", paramsArray);
				NetOffice.ExcelApi.PivotField newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotField;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotItem PivotItem
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotItem", paramsArray);
				NetOffice.ExcelApi.PivotItem newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PivotItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotItem;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotTable", paramsArray);
				NetOffice.ExcelApi.PivotTable newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotTable;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Precedents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Precedents", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Previous
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Previous", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.QueryTable QueryTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QueryTable", paramsArray);
				NetOffice.ExcelApi.QueryTable newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.QueryTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.QueryTable;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1, object cell2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1, cell2);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Resize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Resize", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShowDetail
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowDetail", paramsArray);
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
				Invoker.PropertySet(this, "ShowDetail", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShrinkToFit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShrinkToFit", paramsArray);
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
				Invoker.PropertySet(this, "ShrinkToFit", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SoundNote SoundNote
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SoundNote", paramsArray);
				NetOffice.ExcelApi.SoundNote newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SoundNote.LateBindingApiWrapperType) as NetOffice.ExcelApi.SoundNote;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
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
				Invoker.PropertySet(this, "Style", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Summary
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Summary", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Validation Validation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Validation", paramsArray);
				NetOffice.ExcelApi.Validation newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Validation.LateBindingApiWrapperType) as NetOffice.ExcelApi.Validation;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Worksheet Worksheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Worksheet", paramsArray);
				NetOffice.ExcelApi.Worksheet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Worksheet.LateBindingApiWrapperType) as NetOffice.ExcelApi.Worksheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object WrapText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WrapText", paramsArray);
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
				Invoker.PropertySet(this, "WrapText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comment Comment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comment", paramsArray);
				NetOffice.ExcelApi.Comment newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Comment.LateBindingApiWrapperType) as NetOffice.ExcelApi.Comment;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Phonetic Phonetic
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Phonetic", paramsArray);
				NetOffice.ExcelApi.Phonetic newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Phonetic.LateBindingApiWrapperType) as NetOffice.ExcelApi.Phonetic;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FormatConditions FormatConditions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormatConditions", paramsArray);
				NetOffice.ExcelApi.FormatConditions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.FormatConditions.LateBindingApiWrapperType) as NetOffice.ExcelApi.FormatConditions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 ReadingOrder
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadingOrder", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReadingOrder", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Hyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.ExcelApi.Hyperlinks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Hyperlinks.LateBindingApiWrapperType) as NetOffice.ExcelApi.Hyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Phonetics Phonetics
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Phonetics", paramsArray);
				NetOffice.ExcelApi.Phonetics newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Phonetics.LateBindingApiWrapperType) as NetOffice.ExcelApi.Phonetics;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCell PivotCell
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotCell", paramsArray);
				NetOffice.ExcelApi.PivotCell newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.PivotCell.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotCell;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Errors Errors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Errors", paramsArray);
				NetOffice.ExcelApi.Errors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Errors.LateBindingApiWrapperType) as NetOffice.ExcelApi.Errors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SmartTags SmartTags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTags", paramsArray);
				NetOffice.ExcelApi.SmartTags newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SmartTags.LateBindingApiWrapperType) as NetOffice.ExcelApi.SmartTags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool AllowEdit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllowEdit", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject ListObject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListObject", paramsArray);
				NetOffice.ExcelApi.ListObject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType) as NetOffice.ExcelApi.ListObject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XPath XPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XPath", paramsArray);
				NetOffice.ExcelApi.XPath newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.XPath.LateBindingApiWrapperType) as NetOffice.ExcelApi.XPath;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Actions ServerActions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerActions", paramsArray);
				NetOffice.ExcelApi.Actions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Actions.LateBindingApiWrapperType) as NetOffice.ExcelApi.Actions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public string MDX
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MDX", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object CountLarge
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CountLarge", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SparklineGroups SparklineGroups
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SparklineGroups", paramsArray);
				NetOffice.ExcelApi.SparklineGroups newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.SparklineGroups.LateBindingApiWrapperType) as NetOffice.ExcelApi.SparklineGroups;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.DisplayFormat DisplayFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayFormat", paramsArray);
				NetOffice.ExcelApi.DisplayFormat newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.DisplayFormat.LateBindingApiWrapperType) as NetOffice.ExcelApi.DisplayFormat;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Activate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Activate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction Action</param>
		/// <param name="criteriaRange">optional object CriteriaRange</param>
		/// <param name="copyToRange">optional object CopyToRange</param>
		/// <param name="unique">optional object Unique</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange, object unique)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, criteriaRange, copyToRange, unique);
			object returnItem = Invoker.MethodReturn(this, "AdvancedFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction Action</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action);
			object returnItem = Invoker.MethodReturn(this, "AdvancedFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction Action</param>
		/// <param name="criteriaRange">optional object CriteriaRange</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, criteriaRange);
			object returnItem = Invoker.MethodReturn(this, "AdvancedFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction Action</param>
		/// <param name="criteriaRange">optional object CriteriaRange</param>
		/// <param name="copyToRange">optional object CopyToRange</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(action, criteriaRange, copyToRange);
			object returnItem = Invoker.MethodReturn(this, "AdvancedFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		/// <param name="ignoreRelativeAbsolute">optional object IgnoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object UseRowColumnNames</param>
		/// <param name="omitColumn">optional object OmitColumn</param>
		/// <param name="omitRow">optional object OmitRow</param>
		/// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
		/// <param name="appendLast">optional object AppendLast</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order, object appendLast)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order, appendLast);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		/// <param name="ignoreRelativeAbsolute">optional object IgnoreRelativeAbsolute</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names, ignoreRelativeAbsolute);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		/// <param name="ignoreRelativeAbsolute">optional object IgnoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object UseRowColumnNames</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names, ignoreRelativeAbsolute, useRowColumnNames);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		/// <param name="ignoreRelativeAbsolute">optional object IgnoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object UseRowColumnNames</param>
		/// <param name="omitColumn">optional object OmitColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		/// <param name="ignoreRelativeAbsolute">optional object IgnoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object UseRowColumnNames</param>
		/// <param name="omitColumn">optional object OmitColumn</param>
		/// <param name="omitRow">optional object OmitRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="names">optional object Names</param>
		/// <param name="ignoreRelativeAbsolute">optional object IgnoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object UseRowColumnNames</param>
		/// <param name="omitColumn">optional object OmitColumn</param>
		/// <param name="omitRow">optional object OmitRow</param>
		/// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order);
			object returnItem = Invoker.MethodReturn(this, "ApplyNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ApplyOutlineStyles()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ApplyOutlineStyles", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string AutoComplete(string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(this, "AutoComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">NetOffice.ExcelApi.Range Destination</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlAutoFillType Type = 0</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFill(NetOffice.ExcelApi.Range destination, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, type);
			object returnItem = Invoker.MethodReturn(this, "AutoFill", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">NetOffice.ExcelApi.Range Destination</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFill(NetOffice.ExcelApi.Range destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			object returnItem = Invoker.MethodReturn(this, "AutoFill", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		/// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
		/// <param name="criteria2">optional object Criteria2</param>
		/// <param name="visibleDropDown">optional object VisibleDropDown</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1, _operator, criteria2, visibleDropDown);
			object returnItem = Invoker.MethodReturn(this, "AutoFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AutoFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field);
			object returnItem = Invoker.MethodReturn(this, "AutoFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1);
			object returnItem = Invoker.MethodReturn(this, "AutoFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		/// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1, object _operator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1, _operator);
			object returnItem = Invoker.MethodReturn(this, "AutoFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="field">optional object Field</param>
		/// <param name="criteria1">optional object Criteria1</param>
		/// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
		/// <param name="criteria2">optional object Criteria2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1, object _operator, object criteria2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(field, criteria1, _operator, criteria2);
			object returnItem = Invoker.MethodReturn(this, "AutoFilter", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFit()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AutoFit", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object Number</param>
		/// <param name="font">optional object Font</param>
		/// <param name="alignment">optional object Alignment</param>
		/// <param name="border">optional object Border</param>
		/// <param name="pattern">optional object Pattern</param>
		/// <param name="width">optional object Width</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment, object border, object pattern, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, number, font, alignment, border, pattern, width);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object Number</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, number);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object Number</param>
		/// <param name="font">optional object Font</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, number, font);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object Number</param>
		/// <param name="font">optional object Font</param>
		/// <param name="alignment">optional object Alignment</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, number, font, alignment);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object Number</param>
		/// <param name="font">optional object Font</param>
		/// <param name="alignment">optional object Alignment</param>
		/// <param name="border">optional object Border</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment, object border)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, number, font, alignment, border);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object Number</param>
		/// <param name="font">optional object Font</param>
		/// <param name="alignment">optional object Alignment</param>
		/// <param name="border">optional object Border</param>
		/// <param name="pattern">optional object Pattern</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment, object border, object pattern)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, number, font, alignment, border, pattern);
			object returnItem = Invoker.MethodReturn(this, "AutoFormat", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object AutoOutline()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AutoOutline", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object Color</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle, object weight, object colorIndex, object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex, color);
			object returnItem = Invoker.MethodReturn(this, "BorderAround", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object Color</param>
		/// <param name="themeColor">optional object ThemeColor</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object BorderAround(object lineStyle, object weight, object colorIndex, object color, object themeColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex, color, themeColor);
			object returnItem = Invoker.MethodReturn(this, "BorderAround", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BorderAround", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle);
			object returnItem = Invoker.MethodReturn(this, "BorderAround", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle, object weight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight);
			object returnItem = Invoker.MethodReturn(this, "BorderAround", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle, object weight, object colorIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex);
			object returnItem = Invoker.MethodReturn(this, "BorderAround", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Calculate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Calculate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		/// <param name="spellLang">optional object SpellLang</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="alwaysSuggest">optional object AlwaysSuggest</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(customDictionary, ignoreUppercase, alwaysSuggest);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Clear()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Clear", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ClearContents()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ClearContents", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ClearFormats()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ClearFormats", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ClearNotes()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ClearNotes", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ClearOutline()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ClearOutline", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="comparison">object Comparison</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range ColumnDifferences(object comparison)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison);
			object returnItem = Invoker.MethodReturn(this, "ColumnDifferences", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sources">optional object Sources</param>
		/// <param name="function">optional object Function</param>
		/// <param name="topRow">optional object TopRow</param>
		/// <param name="leftColumn">optional object LeftColumn</param>
		/// <param name="createLinks">optional object CreateLinks</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function, object topRow, object leftColumn, object createLinks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sources, function, topRow, leftColumn, createLinks);
			object returnItem = Invoker.MethodReturn(this, "Consolidate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Consolidate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sources">optional object Sources</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sources);
			object returnItem = Invoker.MethodReturn(this, "Consolidate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sources">optional object Sources</param>
		/// <param name="function">optional object Function</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sources, function);
			object returnItem = Invoker.MethodReturn(this, "Consolidate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sources">optional object Sources</param>
		/// <param name="function">optional object Function</param>
		/// <param name="topRow">optional object TopRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function, object topRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sources, function, topRow);
			object returnItem = Invoker.MethodReturn(this, "Consolidate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sources">optional object Sources</param>
		/// <param name="function">optional object Function</param>
		/// <param name="topRow">optional object TopRow</param>
		/// <param name="leftColumn">optional object LeftColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function, object topRow, object leftColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sources, function, topRow, leftColumn);
			object returnItem = Invoker.MethodReturn(this, "Consolidate", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Copy(object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			object returnItem = Invoker.MethodReturn(this, "Copy", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Copy()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Copy", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="data">object Data</param>
		/// <param name="maxRows">optional object MaxRows</param>
		/// <param name="maxColumns">optional object MaxColumns</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 CopyFromRecordset(object data, object maxRows, object maxColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data, maxRows, maxColumns);
			object returnItem = Invoker.MethodReturn(this, "CopyFromRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="data">object Data</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 CopyFromRecordset(object data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data);
			object returnItem = Invoker.MethodReturn(this, "CopyFromRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="data">object Data</param>
		/// <param name="maxRows">optional object MaxRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 CopyFromRecordset(object data, object maxRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data, maxRows);
			object returnItem = Invoker.MethodReturn(this, "CopyFromRecordset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlCopyPictureFormat Format = -4147</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CopyPicture(object appearance, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance, format);
			object returnItem = Invoker.MethodReturn(this, "CopyPicture", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CopyPicture()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CopyPicture", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CopyPicture(object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(appearance);
			object returnItem = Invoker.MethodReturn(this, "CopyPicture", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="top">optional object Top</param>
		/// <param name="left">optional object Left</param>
		/// <param name="bottom">optional object Bottom</param>
		/// <param name="right">optional object Right</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top, object left, object bottom, object right)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(top, left, bottom, right);
			object returnItem = Invoker.MethodReturn(this, "CreateNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="top">optional object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(top);
			object returnItem = Invoker.MethodReturn(this, "CreateNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="top">optional object Top</param>
		/// <param name="left">optional object Left</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(top, left);
			object returnItem = Invoker.MethodReturn(this, "CreateNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="top">optional object Top</param>
		/// <param name="left">optional object Left</param>
		/// <param name="bottom">optional object Bottom</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top, object left, object bottom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(top, left, bottom);
			object returnItem = Invoker.MethodReturn(this, "CreateNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsBIFF">optional object ContainsBIFF</param>
		/// <param name="containsRTF">optional object ContainsRTF</param>
		/// <param name="containsVALU">optional object ContainsVALU</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF, object containsVALU)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, containsPICT, containsBIFF, containsRTF, containsVALU);
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition);
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance);
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, containsPICT);
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsBIFF">optional object ContainsBIFF</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, containsPICT, containsBIFF);
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">optional object Edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object ContainsPICT</param>
		/// <param name="containsBIFF">optional object ContainsBIFF</param>
		/// <param name="containsRTF">optional object ContainsRTF</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, appearance, containsPICT, containsBIFF, containsRTF);
			object returnItem = Invoker.MethodReturn(this, "CreatePublisher", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Cut(object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			object returnItem = Invoker.MethodReturn(this, "Cut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Cut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Cut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowcol">optional object Rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		/// <param name="step">optional object Step</param>
		/// <param name="stop">optional object Stop</param>
		/// <param name="trend">optional object Trend</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date, object step, object stop, object trend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowcol, type, date, step, stop, trend);
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowcol">optional object Rowcol</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowcol);
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowcol">optional object Rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowcol, type);
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowcol">optional object Rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowcol, type, date);
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowcol">optional object Rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		/// <param name="step">optional object Step</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date, object step)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowcol, type, date, step);
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowcol">optional object Rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		/// <param name="step">optional object Step</param>
		/// <param name="stop">optional object Stop</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date, object step, object stop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowcol, type, date, step, stop);
			object returnItem = Invoker.MethodReturn(this, "DataSeries", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shift">optional object Shift</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Delete(object shift)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shift);
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Delete()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Delete", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object DialogBox()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DialogBox", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType Type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption Option</param>
		/// <param name="name">optional object Name</param>
		/// <param name="reference">optional object Reference</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
		/// <param name="format">optional object Format</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name, reference, appearance, chartSize, format);
			object returnItem = Invoker.MethodReturn(this, "EditionOptions", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType Type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption Option</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option);
			object returnItem = Invoker.MethodReturn(this, "EditionOptions", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType Type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption Option</param>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name);
			object returnItem = Invoker.MethodReturn(this, "EditionOptions", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType Type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption Option</param>
		/// <param name="name">optional object Name</param>
		/// <param name="reference">optional object Reference</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name, reference);
			object returnItem = Invoker.MethodReturn(this, "EditionOptions", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType Type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption Option</param>
		/// <param name="name">optional object Name</param>
		/// <param name="reference">optional object Reference</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name, reference, appearance);
			object returnItem = Invoker.MethodReturn(this, "EditionOptions", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType Type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption Option</param>
		/// <param name="name">optional object Name</param>
		/// <param name="reference">optional object Reference</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, option, name, reference, appearance, chartSize);
			object returnItem = Invoker.MethodReturn(this, "EditionOptions", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FillDown()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FillDown", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FillLeft()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FillLeft", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FillRight()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FillRight", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FillUp()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FillUp", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="matchByte">optional object MatchByte</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="matchByte">optional object MatchByte</param>
		/// <param name="searchFormat">optional object SearchFormat</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte, object searchFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, searchFormat);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="after">optional object After</param>
		/// <param name="lookIn">optional object LookIn</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object MatchCase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Find", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindNext(object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after);
			object returnItem = Invoker.MethodReturn(this, "FindNext", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindNext()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindNext", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindPrevious(object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after);
			object returnItem = Invoker.MethodReturn(this, "FindPrevious", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindPrevious()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindPrevious", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object FunctionWizard()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FunctionWizard", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="goal">object Goal</param>
		/// <param name="changingCell">NetOffice.ExcelApi.Range ChangingCell</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool GoalSeek(object goal, NetOffice.ExcelApi.Range changingCell)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(goal, changingCell);
			object returnItem = Invoker.MethodReturn(this, "GoalSeek", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">optional object Start</param>
		/// <param name="end">optional object End</param>
		/// <param name="by">optional object By</param>
		/// <param name="periods">optional object Periods</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start, object end, object by, object periods)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, end, by, periods);
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Group()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">optional object Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start);
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">optional object Start</param>
		/// <param name="end">optional object End</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start, object end)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, end);
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="start">optional object Start</param>
		/// <param name="end">optional object End</param>
		/// <param name="by">optional object By</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start, object end, object by)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(start, end, by);
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="insertAmount">Int32 InsertAmount</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 InsertIndent(Int32 insertAmount)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(insertAmount);
			object returnItem = Invoker.MethodReturn(this, "InsertIndent", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shift">optional object Shift</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Insert(object shift)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shift);
			object returnItem = Invoker.MethodReturn(this, "Insert", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shift">optional object Shift</param>
		/// <param name="copyOrigin">optional object CopyOrigin</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Insert(object shift, object copyOrigin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shift, copyOrigin);
			object returnItem = Invoker.MethodReturn(this, "Insert", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Insert()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Insert", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Justify()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Justify", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ListNames()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ListNames", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="across">optional object Across</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Merge(object across)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(across);
			object returnItem = Invoker.MethodReturn(this, "Merge", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Merge()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Merge", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 UnMerge()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "UnMerge", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="towardPrecedent">optional object TowardPrecedent</param>
		/// <param name="arrowNumber">optional object ArrowNumber</param>
		/// <param name="linkNumber">optional object LinkNumber</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow(object towardPrecedent, object arrowNumber, object linkNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(towardPrecedent, arrowNumber, linkNumber);
			object returnItem = Invoker.MethodReturn(this, "NavigateArrow", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NavigateArrow", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="towardPrecedent">optional object TowardPrecedent</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow(object towardPrecedent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(towardPrecedent);
			object returnItem = Invoker.MethodReturn(this, "NavigateArrow", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="towardPrecedent">optional object TowardPrecedent</param>
		/// <param name="arrowNumber">optional object ArrowNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow(object towardPrecedent, object arrowNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(towardPrecedent, arrowNumber);
			object returnItem = Invoker.MethodReturn(this, "NavigateArrow", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="text">optional object Text</param>
		/// <param name="start">optional object Start</param>
		/// <param name="length">optional object Length</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string NoteText(object text, object start, object length)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, start, length);
			object returnItem = Invoker.MethodReturn(this, "NoteText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string NoteText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NoteText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="text">optional object Text</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string NoteText(object text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			object returnItem = Invoker.MethodReturn(this, "NoteText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="text">optional object Text</param>
		/// <param name="start">optional object Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public string NoteText(object text, object start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, start);
			object returnItem = Invoker.MethodReturn(this, "NoteText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="parseLine">optional object ParseLine</param>
		/// <param name="destination">optional object Destination</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Parse(object parseLine, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parseLine, destination);
			object returnItem = Invoker.MethodReturn(this, "Parse", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Parse()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Parse", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="parseLine">optional object ParseLine</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Parse(object parseLine)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parseLine);
			object returnItem = Invoker.MethodReturn(this, "Parse", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object SkipBlanks</param>
		/// <param name="transpose">optional object Transpose</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste, object operation, object skipBlanks, object transpose)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste, operation, skipBlanks, transpose);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste, object operation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste, operation);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object SkipBlanks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste, object operation, object skipBlanks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste, operation, skipBlanks);
			object returnItem = Invoker.MethodReturn(this, "PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="prToFileName">optional object PrToFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile);
			object returnItem = Invoker.MethodReturn(this, "_PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="enableChanges">optional object EnableChanges</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintPreview(object enableChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(enableChanges);
			object returnItem = Invoker.MethodReturn(this, "PrintPreview", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintPreview()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PrintPreview", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object RemoveSubtotal()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "RemoveSubtotal", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="matchByte">optional object MatchByte</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement, lookAt, searchOrder, matchCase, matchByte);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="matchByte">optional object MatchByte</param>
		/// <param name="searchFormat">optional object SearchFormat</param>
		/// <param name="replaceFormat">optional object ReplaceFormat</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat, object replaceFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat, replaceFormat);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		/// <param name="lookAt">optional object LookAt</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement, lookAt);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement, lookAt, searchOrder);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="matchCase">optional object MatchCase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement, lookAt, searchOrder, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="what">object What</param>
		/// <param name="replacement">object Replacement</param>
		/// <param name="lookAt">optional object LookAt</param>
		/// <param name="searchOrder">optional object SearchOrder</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="matchByte">optional object MatchByte</param>
		/// <param name="searchFormat">optional object SearchFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat);
			object returnItem = Invoker.MethodReturn(this, "Replace", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="comparison">object Comparison</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range RowDifferences(object comparison)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(comparison);
			object returnItem = Invoker.MethodReturn(this, "RowDifferences", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		/// <param name="arg30">optional object Arg30</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="arg1">optional object Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		/// <param name="arg5">optional object Arg5</param>
		/// <param name="arg6">optional object Arg6</param>
		/// <param name="arg7">optional object Arg7</param>
		/// <param name="arg8">optional object Arg8</param>
		/// <param name="arg9">optional object Arg9</param>
		/// <param name="arg10">optional object Arg10</param>
		/// <param name="arg11">optional object Arg11</param>
		/// <param name="arg12">optional object Arg12</param>
		/// <param name="arg13">optional object Arg13</param>
		/// <param name="arg14">optional object Arg14</param>
		/// <param name="arg15">optional object Arg15</param>
		/// <param name="arg16">optional object Arg16</param>
		/// <param name="arg17">optional object Arg17</param>
		/// <param name="arg18">optional object Arg18</param>
		/// <param name="arg19">optional object Arg19</param>
		/// <param name="arg20">optional object Arg20</param>
		/// <param name="arg21">optional object Arg21</param>
		/// <param name="arg22">optional object Arg22</param>
		/// <param name="arg23">optional object Arg23</param>
		/// <param name="arg24">optional object Arg24</param>
		/// <param name="arg25">optional object Arg25</param>
		/// <param name="arg26">optional object Arg26</param>
		/// <param name="arg27">optional object Arg27</param>
		/// <param name="arg28">optional object Arg28</param>
		/// <param name="arg29">optional object Arg29</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Select()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Select", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Show()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Show", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="remove">optional object Remove</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShowDependents(object remove)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(remove);
			object returnItem = Invoker.MethodReturn(this, "ShowDependents", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShowDependents()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ShowDependents", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShowErrors()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ShowErrors", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="remove">optional object Remove</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShowPrecedents(object remove)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(remove);
			object returnItem = Invoker.MethodReturn(this, "ShowPrecedents", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object ShowPrecedents()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ShowPrecedents", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		/// <param name="dataOption3">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption3 = 0</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2, object dataOption3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, dataOption3);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="type">optional object Type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2);
			object returnItem = Invoker.MethodReturn(this, "Sort", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		/// <param name="dataOption3">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption3 = 0</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2, object dataOption3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2, dataOption3);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object Key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object Type</param>
		/// <param name="key2">optional object Key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object Key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object OrderCustom</param>
		/// <param name="matchCase">optional object MatchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2);
			object returnItem = Invoker.MethodReturn(this, "SortSpecial", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlCellType Type</param>
		/// <param name="value">optional object Value</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, value);
			object returnItem = Invoker.MethodReturn(this, "SpecialCells", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlCellType Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "SpecialCells", paramsArray);
			NetOffice.ExcelApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Range.LateBindingApiWrapperType) as NetOffice.ExcelApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">string Edition</param>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlSubscribeToFormat Format = -4158</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SubscribeTo(string edition, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition, format);
			object returnItem = Invoker.MethodReturn(this, "SubscribeTo", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="edition">string Edition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object SubscribeTo(string edition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(edition);
			object returnItem = Invoker.MethodReturn(this, "SubscribeTo", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="groupBy">Int32 GroupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction Function</param>
		/// <param name="totalList">object TotalList</param>
		/// <param name="replace">optional object Replace</param>
		/// <param name="pageBreaks">optional object PageBreaks</param>
		/// <param name="summaryBelowData">optional NetOffice.ExcelApi.Enums.XlSummaryRow SummaryBelowData = 1</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks, object summaryBelowData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupBy, function, totalList, replace, pageBreaks, summaryBelowData);
			object returnItem = Invoker.MethodReturn(this, "Subtotal", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="groupBy">Int32 GroupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction Function</param>
		/// <param name="totalList">object TotalList</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupBy, function, totalList);
			object returnItem = Invoker.MethodReturn(this, "Subtotal", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="groupBy">Int32 GroupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction Function</param>
		/// <param name="totalList">object TotalList</param>
		/// <param name="replace">optional object Replace</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupBy, function, totalList, replace);
			object returnItem = Invoker.MethodReturn(this, "Subtotal", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="groupBy">Int32 GroupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction Function</param>
		/// <param name="totalList">object TotalList</param>
		/// <param name="replace">optional object Replace</param>
		/// <param name="pageBreaks">optional object PageBreaks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(groupBy, function, totalList, replace, pageBreaks);
			object returnItem = Invoker.MethodReturn(this, "Subtotal", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowInput">optional object RowInput</param>
		/// <param name="columnInput">optional object ColumnInput</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Table(object rowInput, object columnInput)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowInput, columnInput);
			object returnItem = Invoker.MethodReturn(this, "Table", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Table()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Table", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="rowInput">optional object RowInput</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Table(object rowInput)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowInput);
			object returnItem = Invoker.MethodReturn(this, "Table", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		/// <param name="thousandsSeparator">optional object ThousandsSeparator</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		/// <param name="thousandsSeparator">optional object ThousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object TrailingMinusNumbers</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator, trailingMinusNumbers);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator);
			object returnItem = Invoker.MethodReturn(this, "TextToColumns", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Ungroup()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Ungroup", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="text">optional object Text</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comment AddComment(object text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			object returnItem = Invoker.MethodReturn(this, "AddComment", paramsArray);
			NetOffice.ExcelApi.Comment newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Comment.LateBindingApiWrapperType) as NetOffice.ExcelApi.Comment;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comment AddComment()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddComment", paramsArray);
			NetOffice.ExcelApi.Comment newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Comment.LateBindingApiWrapperType) as NetOffice.ExcelApi.Comment;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 ClearComments()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ClearComments", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SetPhonetic()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "SetPhonetic", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="prToFileName">optional object PrToFileName</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate, prToFileName);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate);
			object returnItem = Invoker.MethodReturn(this, "PrintOut", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object SkipBlanks</param>
		/// <param name="transpose">optional object Transpose</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste, object operation, object skipBlanks, object transpose)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste, operation, skipBlanks, transpose);
			object returnItem = Invoker.MethodReturn(this, "_PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste);
			object returnItem = Invoker.MethodReturn(this, "_PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste, object operation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste, operation);
			object returnItem = Invoker.MethodReturn(this, "_PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object SkipBlanks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste, object operation, object skipBlanks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(paste, operation, skipBlanks);
			object returnItem = Invoker.MethodReturn(this, "_PasteSpecial", paramsArray);
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 Dirty()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Dirty", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="speakDirection">optional object SpeakDirection</param>
		/// <param name="speakFormulas">optional object SpeakFormulas</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 Speak(object speakDirection, object speakFormulas)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(speakDirection, speakFormulas);
			object returnItem = Invoker.MethodReturn(this, "Speak", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 Speak()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Speak", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="speakDirection">optional object SpeakDirection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 Speak(object speakDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(speakDirection);
			object returnItem = Invoker.MethodReturn(this, "Speak", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile, collate);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="preview">optional object Preview</param>
		/// <param name="activePrinter">optional object ActivePrinter</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, copies, preview, activePrinter, printToFile);
			object returnItem = Invoker.MethodReturn(this, "__PrintOut", paramsArray);
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="columns">optional object Columns</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 RemoveDuplicates(object columns, object header)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(columns, header);
			object returnItem = Invoker.MethodReturn(this, "RemoveDuplicates", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 RemoveDuplicates()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "RemoveDuplicates", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="columns">optional object Columns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 RemoveDuplicates(object columns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(columns);
			object returnItem = Invoker.MethodReturn(this, "RemoveDuplicates", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="openAfterPublish">optional object OpenAfterPublish</param>
		/// <param name="fixedFormatExtClassPtr">optional object FixedFormatExtClassPtr</param>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType Type</param>
		/// <param name="filename">optional object Filename</param>
		/// <param name="quality">optional object Quality</param>
		/// <param name="includeDocProperties">optional object IncludeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object IgnorePrintAreas</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="openAfterPublish">optional object OpenAfterPublish</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish);
			object returnItem = Invoker.MethodReturn(this, "ExportAsFixedFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 12,14,15,16)]
		public object CalculateRowMajorOrder()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CalculateRowMajorOrder", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object Color</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle, object weight, object colorIndex, object color)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex, color);
			object returnItem = Invoker.MethodReturn(this, "_BorderAround", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object _BorderAround()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "_BorderAround", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle);
			object returnItem = Invoker.MethodReturn(this, "_BorderAround", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle, object weight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight);
			object returnItem = Invoker.MethodReturn(this, "_BorderAround", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lineStyle">optional object LineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle, object weight, object colorIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineStyle, weight, colorIndex);
			object returnItem = Invoker.MethodReturn(this, "_BorderAround", paramsArray);
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
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 ClearHyperlinks()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ClearHyperlinks", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 AllocateChanges()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AllocateChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 DiscardChanges()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DiscardChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 FlashFill()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FlashFill", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

       #region IEnumerable<object> Member
        
        /// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
       public IEnumerator<object> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (object item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}