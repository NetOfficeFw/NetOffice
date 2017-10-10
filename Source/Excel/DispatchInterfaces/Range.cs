using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Range
	/// </summary>
	[SyntaxBypass]
 	public class Range_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Range_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Range_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// <param name="relativeTo">optional object relativeTo</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", new object[]{ rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// <param name="relativeTo">optional object relativeTo</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Address")]
		public string Address(object rowAbsolute)
		{
			return get_Address(rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute)
		{
			return get_Address(rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return Factory.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Address
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Address")]
		public string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// <param name="relativeTo">optional object relativeTo</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return Factory.ExecuteStringPropertyGet(this, "AddressLocal", new object[]{ rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// <param name="relativeTo">optional object relativeTo</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_AddressLocal")]
		public string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute)
		{
			return Factory.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_AddressLocal")]
		public string AddressLocal(object rowAbsolute)
		{
			return get_AddressLocal(rowAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute)
		{
			return Factory.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_AddressLocal")]
		public string AddressLocal(object rowAbsolute, object columnAbsolute)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return Factory.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_AddressLocal")]
		public string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return Factory.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_AddressLocal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
		/// <param name="rowAbsolute">optional object rowAbsolute</param>
		/// <param name="columnAbsolute">optional object columnAbsolute</param>
		/// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
		/// <param name="external">optional object external</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_AddressLocal")]
		public string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
		{
			return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional object start</param>
		/// <param name="length">optional object length</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Characters get_Characters(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", NetOffice.ExcelApi.Characters.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Characters
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Characters")]
		public NetOffice.ExcelApi.Characters Characters(object start, object length)
		{
			return get_Characters(start, length);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional object start</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Characters get_Characters(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", NetOffice.ExcelApi.Characters.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Characters
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx </remarks>
		/// <param name="start">optional object start</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Characters")]
		public NetOffice.ExcelApi.Characters Characters(object start)
		{
			return get_Characters(start);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object rowOffset</param>
		/// <param name="columnOffset">optional object columnOffset</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Offset(object rowOffset, object columnOffset)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Offset", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, rowOffset, columnOffset);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Offset
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx </remarks>
		/// <param name="rowOffset">optional object rowOffset</param>
		/// <param name="columnOffset">optional object columnOffset</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Offset")]
		public NetOffice.ExcelApi.Range Offset(object rowOffset, object columnOffset)
		{
			return get_Offset(rowOffset, columnOffset);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowOffset">optional object rowOffset</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Offset(object rowOffset)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Offset", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, rowOffset);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Offset
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx </remarks>
		/// <param name="rowOffset">optional object rowOffset</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Offset")]
		public NetOffice.ExcelApi.Range Offset(object rowOffset)
		{
			return get_Offset(rowOffset);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1, object cell2)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx </remarks>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public NetOffice.ExcelApi.Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834676.aspx </remarks>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public NetOffice.ExcelApi.Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowSize">optional object rowSize</param>
		/// <param name="columnSize">optional object columnSize</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Resize(object rowSize, object columnSize)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Resize", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, rowSize, columnSize);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Resize
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx </remarks>
		/// <param name="rowSize">optional object rowSize</param>
		/// <param name="columnSize">optional object columnSize</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Resize")]
		public NetOffice.ExcelApi.Range Resize(object rowSize, object columnSize)
		{
			return get_Resize(rowSize, columnSize);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="rowSize">optional object rowSize</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Resize(object rowSize)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Resize", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, rowSize);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Resize
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx </remarks>
		/// <param name="rowSize">optional object rowSize</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Resize")]
		public NetOffice.ExcelApi.Range Resize(object rowSize)
		{
			return get_Resize(rowSize);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="rangeValueDataType">optional object rangeValueDataType</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195193.aspx
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Value(object rangeValueDataType)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Value", rangeValueDataType);
		}

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Value(object rangeValueDataType, object value)
		{
			Factory.ExecutePropertySet(this, "Value", rangeValueDataType, value);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Alias for get_Value
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195193.aspx </remarks>
		/// <param name="rangeValueDataType">optional object rangeValueDataType</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16), Redirect("get_Value")]
		public object Value(object rangeValueDataType)
		{
			return get_Value(rangeValueDataType);
		}

		#endregion
	}

	/// <summary>
	/// DispatchInterface Range 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838238.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class Range : Range_ , IEnumerableProvider<NetOffice.ExcelApi.Range>
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
                    _type = typeof(Range);
                return _type;
            }
        }

        #endregion

        #region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Range(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

        ///<param name="factory">current used factory core</param>
        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        public Range(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193954.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839656.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196885.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197727.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AddIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AddIndent");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AddIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837625.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Address");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194992.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string AddressLocal
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AddressLocal");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196243.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Areas Areas
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Areas>(this, "Areas", NetOffice.ExcelApi.Areas.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822605.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Borders Borders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Borders>(this, "Borders", NetOffice.ExcelApi.Borders.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196273.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Cells
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Cells", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198232.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Characters Characters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", NetOffice.ExcelApi.Characters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198200.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Column
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Column");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837125.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Columns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Columns", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837430.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193349.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194268.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range CurrentArray
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "CurrentArray", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196678.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range CurrentRegion
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "CurrentRegion", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// Custom Indexer
        /// </summary>
        /// <param name="rowIndex">optional object rowIndex</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public NetOffice.ExcelApi.Range this[object rowIndex]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "_Default", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, rowIndex);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "_Default", value, rowIndex);
			}
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rowIndex">optional object rowIndex</param>
        /// <param name="columnIndex">optional object columnIndex</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.Range this[object rowIndex, object columnIndex]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "_Default", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, rowIndex, columnIndex);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "_Default", value, rowIndex, columnIndex);
			}
		}
        
		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197707.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Dependents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Dependents", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195386.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DirectDependents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DirectDependents", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839667.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DirectPrecedents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DirectPrecedents", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839539.aspx </remarks>
		/// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection direction</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_End(NetOffice.ExcelApi.Enums.XlDirection direction)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "End", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, direction);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_End
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839539.aspx </remarks>
		/// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection direction</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_End")]
		public NetOffice.ExcelApi.Range End(NetOffice.ExcelApi.Enums.XlDirection direction)
		{
			return get_End(direction);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834476.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range EntireColumn
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "EntireColumn", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836836.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range EntireRow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "EntireRow", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839770.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Font Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Font>(this, "Font", NetOffice.ExcelApi.Font.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838835.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837104.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlFormulaLabel FormulaLabel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlFormulaLabel>(this, "FormulaLabel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FormulaLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838209.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FormulaHidden
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FormulaHidden");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FormulaHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838851.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823188.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FormulaR1C1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FormulaR1C1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FormulaR1C1", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838568.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FormulaR1C1Local
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FormulaR1C1Local");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FormulaR1C1Local", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841166.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object HasArray
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HasArray");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837123.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object HasFormula
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HasFormula");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840142.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Height
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834657.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Hidden
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Hidden");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Hidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822120.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840976.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object IndentLevel
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IndentLevel");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "IndentLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836210.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Interior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Interior>(this, "Interior", NetOffice.ExcelApi.Interior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821859.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Left
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839644.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 ListHeaderRows
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListHeaderRows");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834387.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlLocationInTable LocationInTable
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlLocationInTable>(this, "LocationInTable");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836172.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822300.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range MergeArea
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "MergeArea", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197310.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196817.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Name
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193940.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Next
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Next", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196401.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840206.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object NumberFormatLocal
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "NumberFormatLocal");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "NumberFormatLocal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840060.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Offset
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Offset", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198186.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Orientation
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Orientation");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838455.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OutlineLevel
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OutlineLevel");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "OutlineLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193644.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PageBreak
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageBreak");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageBreak", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820916.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField PivotField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "PivotField", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193001.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotItem PivotItem
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotItem>(this, "PivotItem", NetOffice.ExcelApi.PivotItem.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837826.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTable", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196936.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Precedents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Precedents", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194949.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrefixCharacter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrefixCharacter");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822759.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Previous
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Previous", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821864.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.QueryTable QueryTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.QueryTable>(this, "QueryTable", NetOffice.ExcelApi.QueryTable.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193274.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Resize
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Resize", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196952.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Row
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Row");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193926.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195745.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Rows", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194526.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowDetail
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ShowDetail");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ShowDetail", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841203.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShrinkToFit
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ShrinkToFit");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ShrinkToFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193277.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SoundNote SoundNote
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SoundNote>(this, "SoundNote", NetOffice.ExcelApi.SoundNote.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834326.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Style
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Style");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841152.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Summary
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Summary");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840217.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Text
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Text");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193765.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Top
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821216.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836425.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839408.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Validation Validation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Validation>(this, "Validation", NetOffice.ExcelApi.Validation.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195193.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193553.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837977.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823126.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Width
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837845.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Worksheet Worksheet
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Worksheet>(this, "Worksheet", NetOffice.ExcelApi.Worksheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821514.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object WrapText
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "WrapText");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "WrapText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836191.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comment Comment
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Comment>(this, "Comment", NetOffice.ExcelApi.Comment.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836776.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Phonetic Phonetic
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Phonetic>(this, "Phonetic", NetOffice.ExcelApi.Phonetic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822167.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FormatConditions FormatConditions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.FormatConditions>(this, "FormatConditions", NetOffice.ExcelApi.FormatConditions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840917.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 ReadingOrder
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ReadingOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839653.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Hyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Hyperlinks>(this, "Hyperlinks", NetOffice.ExcelApi.Hyperlinks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841233.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Phonetics Phonetics
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Phonetics>(this, "Phonetics", NetOffice.ExcelApi.Phonetics.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193914.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string ID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836438.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCell PivotCell
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotCell>(this, "PivotCell", NetOffice.ExcelApi.PivotCell.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835287.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Errors Errors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Errors>(this, "Errors", NetOffice.ExcelApi.Errors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SmartTags SmartTags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SmartTags>(this, "SmartTags", NetOffice.ExcelApi.SmartTags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837052.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowEdit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEdit");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838413.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject ListObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObject>(this, "ListObject", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835889.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XPath XPath
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XPath>(this, "XPath", NetOffice.ExcelApi.XPath.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840072.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Actions ServerActions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Actions>(this, "ServerActions", NetOffice.ExcelApi.Actions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822495.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string MDX
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MDX");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196838.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object CountLarge
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CountLarge");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822140.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SparklineGroups SparklineGroups
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SparklineGroups>(this, "SparklineGroups", NetOffice.ExcelApi.SparklineGroups.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838814.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.DisplayFormat DisplayFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DisplayFormat>(this, "DisplayFormat", NetOffice.ExcelApi.DisplayFormat.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837085.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Activate()
		{
			return Factory.ExecuteVariantMethodGet(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
		/// <param name="criteriaRange">optional object criteriaRange</param>
		/// <param name="copyToRange">optional object copyToRange</param>
		/// <param name="unique">optional object unique</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange, object unique)
		{
			return Factory.ExecuteVariantMethodGet(this, "AdvancedFilter", action, criteriaRange, copyToRange, unique);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action)
		{
			return Factory.ExecuteVariantMethodGet(this, "AdvancedFilter", action);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
		/// <param name="criteriaRange">optional object criteriaRange</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange)
		{
			return Factory.ExecuteVariantMethodGet(this, "AdvancedFilter", action, criteriaRange);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841242.aspx </remarks>
		/// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
		/// <param name="criteriaRange">optional object criteriaRange</param>
		/// <param name="copyToRange">optional object copyToRange</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange)
		{
			return Factory.ExecuteVariantMethodGet(this, "AdvancedFilter", action, criteriaRange, copyToRange);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		/// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object useRowColumnNames</param>
		/// <param name="omitColumn">optional object omitColumn</param>
		/// <param name="omitRow">optional object omitRow</param>
		/// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
		/// <param name="appendLast">optional object appendLast</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order, object appendLast)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", new object[]{ names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order, appendLast });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames()
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", names);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		/// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", names, ignoreRelativeAbsolute);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		/// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object useRowColumnNames</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", names, ignoreRelativeAbsolute, useRowColumnNames);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		/// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object useRowColumnNames</param>
		/// <param name="omitColumn">optional object omitColumn</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		/// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object useRowColumnNames</param>
		/// <param name="omitColumn">optional object omitColumn</param>
		/// <param name="omitRow">optional object omitRow</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", new object[]{ names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196578.aspx </remarks>
		/// <param name="names">optional object names</param>
		/// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
		/// <param name="useRowColumnNames">optional object useRowColumnNames</param>
		/// <param name="omitColumn">optional object omitColumn</param>
		/// <param name="omitRow">optional object omitRow</param>
		/// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyNames", new object[]{ names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840485.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ApplyOutlineStyles()
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyOutlineStyles");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822899.aspx </remarks>
		/// <param name="_string">string string</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string AutoComplete(string _string)
		{
			return Factory.ExecuteStringMethodGet(this, "AutoComplete", _string);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195345.aspx </remarks>
		/// <param name="destination">NetOffice.ExcelApi.Range destination</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlAutoFillType Type = 0</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFill(NetOffice.ExcelApi.Range destination, object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFill", destination, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195345.aspx </remarks>
		/// <param name="destination">NetOffice.ExcelApi.Range destination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFill(NetOffice.ExcelApi.Range destination)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFill", destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		/// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
		/// <param name="criteria2">optional object criteria2</param>
		/// <param name="visibleDropDown">optional object visibleDropDown</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFilter", new object[]{ field, criteria1, _operator, criteria2, visibleDropDown });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter()
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFilter");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
		/// <param name="field">optional object field</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFilter", field);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFilter", field, criteria1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		/// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1, object _operator)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFilter", field, criteria1, _operator);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193884.aspx </remarks>
		/// <param name="field">optional object field</param>
		/// <param name="criteria1">optional object criteria1</param>
		/// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
		/// <param name="criteria2">optional object criteria2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFilter(object field, object criteria1, object _operator, object criteria2)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFilter", field, criteria1, _operator, criteria2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820840.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFit()
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFit");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object number</param>
		/// <param name="font">optional object font</param>
		/// <param name="alignment">optional object alignment</param>
		/// <param name="border">optional object border</param>
		/// <param name="pattern">optional object pattern</param>
		/// <param name="width">optional object width</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment, object border, object pattern, object width)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", new object[]{ format, number, font, alignment, border, pattern, width });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat()
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object number</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", format, number);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object number</param>
		/// <param name="font">optional object font</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", format, number, font);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object number</param>
		/// <param name="font">optional object font</param>
		/// <param name="alignment">optional object alignment</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", format, number, font, alignment);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object number</param>
		/// <param name="font">optional object font</param>
		/// <param name="alignment">optional object alignment</param>
		/// <param name="border">optional object border</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment, object border)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", new object[]{ format, number, font, alignment, border });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
		/// <param name="number">optional object number</param>
		/// <param name="font">optional object font</param>
		/// <param name="alignment">optional object alignment</param>
		/// <param name="border">optional object border</param>
		/// <param name="pattern">optional object pattern</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoFormat(object format, object number, object font, object alignment, object border, object pattern)
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoFormat", new object[]{ format, number, font, alignment, border, pattern });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837148.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object AutoOutline()
		{
			return Factory.ExecuteVariantMethodGet(this, "AutoOutline");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object color</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle, object weight, object colorIndex, object color)
		{
			return Factory.ExecuteVariantMethodGet(this, "BorderAround", lineStyle, weight, colorIndex, color);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object color</param>
		/// <param name="themeColor">optional object themeColor</param>
		[SupportByVersion("Excel", 14,15,16)]
		public object BorderAround(object lineStyle, object weight, object colorIndex, object color, object themeColor)
		{
			return Factory.ExecuteVariantMethodGet(this, "BorderAround", new object[]{ lineStyle, weight, colorIndex, color, themeColor });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround()
		{
			return Factory.ExecuteVariantMethodGet(this, "BorderAround");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
		/// <param name="lineStyle">optional object lineStyle</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle)
		{
			return Factory.ExecuteVariantMethodGet(this, "BorderAround", lineStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle, object weight)
		{
			return Factory.ExecuteVariantMethodGet(this, "BorderAround", lineStyle, weight);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197210.aspx </remarks>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object BorderAround(object lineStyle, object weight, object colorIndex)
		{
			return Factory.ExecuteVariantMethodGet(this, "BorderAround", lineStyle, weight, colorIndex);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834613.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Calculate()
		{
			return Factory.ExecuteVariantMethodGet(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling()
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary)
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195092.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820947.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Clear()
		{
			return Factory.ExecuteVariantMethodGet(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835589.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ClearContents()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearContents");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196576.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ClearFormats()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195197.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ClearNotes()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearNotes");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834886.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ClearOutline()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearOutline");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197752.aspx </remarks>
		/// <param name="comparison">object comparison</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range ColumnDifferences(object comparison)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "ColumnDifferences", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, comparison);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
		/// <param name="sources">optional object sources</param>
		/// <param name="function">optional object function</param>
		/// <param name="topRow">optional object topRow</param>
		/// <param name="leftColumn">optional object leftColumn</param>
		/// <param name="createLinks">optional object createLinks</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function, object topRow, object leftColumn, object createLinks)
		{
			return Factory.ExecuteVariantMethodGet(this, "Consolidate", new object[]{ sources, function, topRow, leftColumn, createLinks });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate()
		{
			return Factory.ExecuteVariantMethodGet(this, "Consolidate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
		/// <param name="sources">optional object sources</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources)
		{
			return Factory.ExecuteVariantMethodGet(this, "Consolidate", sources);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
		/// <param name="sources">optional object sources</param>
		/// <param name="function">optional object function</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function)
		{
			return Factory.ExecuteVariantMethodGet(this, "Consolidate", sources, function);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
		/// <param name="sources">optional object sources</param>
		/// <param name="function">optional object function</param>
		/// <param name="topRow">optional object topRow</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function, object topRow)
		{
			return Factory.ExecuteVariantMethodGet(this, "Consolidate", sources, function, topRow);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839594.aspx </remarks>
		/// <param name="sources">optional object sources</param>
		/// <param name="function">optional object function</param>
		/// <param name="topRow">optional object topRow</param>
		/// <param name="leftColumn">optional object leftColumn</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Consolidate(object sources, object function, object topRow, object leftColumn)
		{
			return Factory.ExecuteVariantMethodGet(this, "Consolidate", sources, function, topRow, leftColumn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837760.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Copy(object destination)
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy", destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837760.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839240.aspx </remarks>
		/// <param name="data">object data</param>
		/// <param name="maxRows">optional object maxRows</param>
		/// <param name="maxColumns">optional object maxColumns</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CopyFromRecordset(object data, object maxRows, object maxColumns)
		{
			return Factory.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows, maxColumns);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839240.aspx </remarks>
		/// <param name="data">object data</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CopyFromRecordset(object data)
		{
			return Factory.ExecuteInt32MethodGet(this, "CopyFromRecordset", data);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839240.aspx </remarks>
		/// <param name="data">object data</param>
		/// <param name="maxRows">optional object maxRows</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CopyFromRecordset(object data, object maxRows)
		{
			return Factory.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193594.aspx </remarks>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlCopyPictureFormat Format = -4147</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CopyPicture(object appearance, object format)
		{
			return Factory.ExecuteVariantMethodGet(this, "CopyPicture", appearance, format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193594.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CopyPicture()
		{
			return Factory.ExecuteVariantMethodGet(this, "CopyPicture");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193594.aspx </remarks>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CopyPicture(object appearance)
		{
			return Factory.ExecuteVariantMethodGet(this, "CopyPicture", appearance);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="bottom">optional object bottom</param>
		/// <param name="right">optional object right</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top, object left, object bottom, object right)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateNames", top, left, bottom, right);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames()
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateNames");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateNames", top);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top, object left)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateNames", top, left);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192960.aspx </remarks>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="bottom">optional object bottom</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreateNames(object top, object left, object bottom)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreateNames", top, left, bottom);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		/// <param name="containsBIFF">optional object containsBIFF</param>
		/// <param name="containsRTF">optional object containsRTF</param>
		/// <param name="containsVALU">optional object containsVALU</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF, object containsVALU)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher", new object[]{ edition, appearance, containsPICT, containsBIFF, containsRTF, containsVALU });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher()
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher", edition);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher", edition, appearance);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher", edition, appearance, containsPICT);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		/// <param name="containsBIFF">optional object containsBIFF</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher", edition, appearance, containsPICT, containsBIFF);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		/// <param name="containsBIFF">optional object containsBIFF</param>
		/// <param name="containsRTF">optional object containsRTF</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF)
		{
			return Factory.ExecuteVariantMethodGet(this, "CreatePublisher", new object[]{ edition, appearance, containsPICT, containsBIFF, containsRTF });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838374.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Cut(object destination)
		{
			return Factory.ExecuteVariantMethodGet(this, "Cut", destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838374.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Cut()
		{
			return Factory.ExecuteVariantMethodGet(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		/// <param name="step">optional object step</param>
		/// <param name="stop">optional object stop</param>
		/// <param name="trend">optional object trend</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date, object step, object stop, object trend)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries", new object[]{ rowcol, type, date, step, stop, trend });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries()
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		/// <param name="rowcol">optional object rowcol</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries", rowcol);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries", rowcol, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries", rowcol, type, date);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		/// <param name="step">optional object step</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date, object step)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries", rowcol, type, date, step);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839262.aspx </remarks>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
		/// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
		/// <param name="step">optional object step</param>
		/// <param name="stop">optional object stop</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DataSeries(object rowcol, object type, object date, object step, object stop)
		{
			return Factory.ExecuteVariantMethodGet(this, "DataSeries", new object[]{ rowcol, type, date, step, stop });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834641.aspx </remarks>
		/// <param name="shift">optional object shift</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Delete(object shift)
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete", shift);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834641.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839446.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DialogBox()
		{
			return Factory.ExecuteVariantMethodGet(this, "DialogBox");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
		/// <param name="name">optional object name</param>
		/// <param name="reference">optional object reference</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize, object format)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditionOptions", new object[]{ type, option, name, reference, appearance, chartSize, format });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditionOptions", type, option);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditionOptions", type, option, name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
		/// <param name="name">optional object name</param>
		/// <param name="reference">optional object reference</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditionOptions", type, option, name, reference);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
		/// <param name="name">optional object name</param>
		/// <param name="reference">optional object reference</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditionOptions", new object[]{ type, option, name, reference, appearance });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821204.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
		/// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
		/// <param name="name">optional object name</param>
		/// <param name="reference">optional object reference</param>
		/// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
		/// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditionOptions", new object[]{ type, option, name, reference, appearance, chartSize });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838404.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FillDown()
		{
			return Factory.ExecuteVariantMethodGet(this, "FillDown");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197298.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FillLeft()
		{
			return Factory.ExecuteVariantMethodGet(this, "FillLeft");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837984.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FillRight()
		{
			return Factory.ExecuteVariantMethodGet(this, "FillRight");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820752.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FillUp()
		{
			return Factory.ExecuteVariantMethodGet(this, "FillUp");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchByte">optional object matchByte</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="searchFormat">optional object searchFormat</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte, object searchFormat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, searchFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, what, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, what, after, lookIn);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, what, after, lookIn, lookAt);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ what, after, lookIn, lookAt, searchOrder });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839746.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="after">optional object after</param>
		/// <param name="lookIn">optional object lookIn</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[]{ what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196143.aspx </remarks>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindNext(object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindNext", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196143.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindNext()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindNext", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838614.aspx </remarks>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindPrevious(object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindPrevious", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838614.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range FindPrevious()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindPrevious", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837608.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FunctionWizard()
		{
			return Factory.ExecuteVariantMethodGet(this, "FunctionWizard");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="goal">object goal</param>
		/// <param name="changingCell">NetOffice.ExcelApi.Range changingCell</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool GoalSeek(object goal, NetOffice.ExcelApi.Range changingCell)
		{
			return Factory.ExecuteBoolMethodGet(this, "GoalSeek", goal, changingCell);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="end">optional object end</param>
		/// <param name="by">optional object by</param>
		/// <param name="periods">optional object periods</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start, object end, object by, object periods)
		{
			return Factory.ExecuteVariantMethodGet(this, "Group", start, end, by, periods);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Group()
		{
			return Factory.ExecuteVariantMethodGet(this, "Group");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start)
		{
			return Factory.ExecuteVariantMethodGet(this, "Group", start);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="end">optional object end</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start, object end)
		{
			return Factory.ExecuteVariantMethodGet(this, "Group", start, end);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839808.aspx </remarks>
		/// <param name="start">optional object start</param>
		/// <param name="end">optional object end</param>
		/// <param name="by">optional object by</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Group(object start, object end, object by)
		{
			return Factory.ExecuteVariantMethodGet(this, "Group", start, end, by);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194874.aspx </remarks>
		/// <param name="insertAmount">Int32 insertAmount</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void InsertIndent(Int32 insertAmount)
		{
			 Factory.ExecuteMethod(this, "InsertIndent", insertAmount);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840310.aspx </remarks>
		/// <param name="shift">optional object shift</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Insert(object shift)
		{
			return Factory.ExecuteVariantMethodGet(this, "Insert", shift);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840310.aspx </remarks>
		/// <param name="shift">optional object shift</param>
		/// <param name="copyOrigin">optional object copyOrigin</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Insert(object shift, object copyOrigin)
		{
			return Factory.ExecuteVariantMethodGet(this, "Insert", shift, copyOrigin);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840310.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Insert()
		{
			return Factory.ExecuteVariantMethodGet(this, "Insert");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841125.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Justify()
		{
			return Factory.ExecuteVariantMethodGet(this, "Justify");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193257.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ListNames()
		{
			return Factory.ExecuteVariantMethodGet(this, "ListNames");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840729.aspx </remarks>
		/// <param name="across">optional object across</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Merge(object across)
		{
			 Factory.ExecuteMethod(this, "Merge", across);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840729.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Merge()
		{
			 Factory.ExecuteMethod(this, "Merge");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840063.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UnMerge()
		{
			 Factory.ExecuteMethod(this, "UnMerge");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
		/// <param name="towardPrecedent">optional object towardPrecedent</param>
		/// <param name="arrowNumber">optional object arrowNumber</param>
		/// <param name="linkNumber">optional object linkNumber</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow(object towardPrecedent, object arrowNumber, object linkNumber)
		{
			return Factory.ExecuteVariantMethodGet(this, "NavigateArrow", towardPrecedent, arrowNumber, linkNumber);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow()
		{
			return Factory.ExecuteVariantMethodGet(this, "NavigateArrow");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
		/// <param name="towardPrecedent">optional object towardPrecedent</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow(object towardPrecedent)
		{
			return Factory.ExecuteVariantMethodGet(this, "NavigateArrow", towardPrecedent);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822883.aspx </remarks>
		/// <param name="towardPrecedent">optional object towardPrecedent</param>
		/// <param name="arrowNumber">optional object arrowNumber</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object NavigateArrow(object towardPrecedent, object arrowNumber)
		{
			return Factory.ExecuteVariantMethodGet(this, "NavigateArrow", towardPrecedent, arrowNumber);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
		/// <param name="text">optional object text</param>
		/// <param name="start">optional object start</param>
		/// <param name="length">optional object length</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string NoteText(object text, object start, object length)
		{
			return Factory.ExecuteStringMethodGet(this, "NoteText", text, start, length);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string NoteText()
		{
			return Factory.ExecuteStringMethodGet(this, "NoteText");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string NoteText(object text)
		{
			return Factory.ExecuteStringMethodGet(this, "NoteText", text);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839198.aspx </remarks>
		/// <param name="text">optional object text</param>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string NoteText(object text, object start)
		{
			return Factory.ExecuteStringMethodGet(this, "NoteText", text, start);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196436.aspx </remarks>
		/// <param name="parseLine">optional object parseLine</param>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Parse(object parseLine, object destination)
		{
			return Factory.ExecuteVariantMethodGet(this, "Parse", parseLine, destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196436.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Parse()
		{
			return Factory.ExecuteVariantMethodGet(this, "Parse");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196436.aspx </remarks>
		/// <param name="parseLine">optional object parseLine</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Parse(object parseLine)
		{
			return Factory.ExecuteVariantMethodGet(this, "Parse", parseLine);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object skipBlanks</param>
		/// <param name="transpose">optional object transpose</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste, object operation, object skipBlanks, object transpose)
		{
			return Factory.ExecuteVariantMethodGet(this, "PasteSpecial", paste, operation, skipBlanks, transpose);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial()
		{
			return Factory.ExecuteVariantMethodGet(this, "PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste)
		{
			return Factory.ExecuteVariantMethodGet(this, "PasteSpecial", paste);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste, object operation)
		{
			return Factory.ExecuteVariantMethodGet(this, "PasteSpecial", paste, operation);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839476.aspx </remarks>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object skipBlanks</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PasteSpecial(object paste, object operation, object skipBlanks)
		{
			return Factory.ExecuteVariantMethodGet(this, "PasteSpecial", paste, operation, skipBlanks);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut()
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838066.aspx </remarks>
		/// <param name="enableChanges">optional object enableChanges</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintPreview(object enableChanges)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintPreview", enableChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838066.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintPreview()
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840556.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object RemoveSubtotal()
		{
			return Factory.ExecuteVariantMethodGet(this, "RemoveSubtotal");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchByte">optional object matchByte</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", new object[]{ what, replacement, lookAt, searchOrder, matchCase, matchByte });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="searchFormat">optional object searchFormat</param>
		/// <param name="replaceFormat">optional object replaceFormat</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat, object replaceFormat)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", new object[]{ what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat, replaceFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", what, replacement);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		/// <param name="lookAt">optional object lookAt</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", what, replacement, lookAt);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", what, replacement, lookAt, searchOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", new object[]{ what, replacement, lookAt, searchOrder, matchCase });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194086.aspx </remarks>
		/// <param name="what">object what</param>
		/// <param name="replacement">object replacement</param>
		/// <param name="lookAt">optional object lookAt</param>
		/// <param name="searchOrder">optional object searchOrder</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="searchFormat">optional object searchFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat)
		{
			return Factory.ExecuteBoolMethodGet(this, "Replace", new object[]{ what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835292.aspx </remarks>
		/// <param name="comparison">object comparison</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range RowDifferences(object comparison)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "RowDifferences", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, comparison);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run()
		{
			return Factory.ExecuteVariantMethodGet(this, "Run");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", arg1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", arg1, arg2, arg3);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", arg1, arg2, arg3, arg4);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838233.aspx </remarks>
		/// <param name="arg1">optional object arg1</param>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197597.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838617.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Show()
		{
			return Factory.ExecuteVariantMethodGet(this, "Show");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840851.aspx </remarks>
		/// <param name="remove">optional object remove</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowDependents(object remove)
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowDependents", remove);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840851.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowDependents()
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowDependents");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192988.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowErrors()
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowErrors");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193004.aspx </remarks>
		/// <param name="remove">optional object remove</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowPrecedents(object remove)
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowPrecedents", remove);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193004.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ShowPrecedents()
		{
			return Factory.ExecuteVariantMethodGet(this, "ShowPrecedents");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		/// <param name="dataOption3">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption3 = 0</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2, object dataOption3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, dataOption3 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort()
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", key1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", key1, order1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", key1, order1, key2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", key1, order1, key2, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840646.aspx </remarks>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="type">optional object type</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Sort", new object[]{ key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		/// <param name="dataOption3">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption3 = 0</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2, object dataOption3)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2, dataOption3 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial()
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod, key1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod, key1, order1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod, key1, order1, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1 });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822807.aspx </remarks>
		/// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
		/// <param name="key1">optional object key1</param>
		/// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
		/// <param name="type">optional object type</param>
		/// <param name="key2">optional object key2</param>
		/// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
		/// <param name="key3">optional object key3</param>
		/// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		/// <param name="orderCustom">optional object orderCustom</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="orientation">optional NetOffice.ExcelApi.Enums.XlSortOrientation Orientation = 2</param>
		/// <param name="dataOption1">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption1 = 0</param>
		/// <param name="dataOption2">optional NetOffice.ExcelApi.Enums.XlSortDataOption DataOption2 = 0</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2)
		{
			return Factory.ExecuteVariantMethodGet(this, "SortSpecial", new object[]{ sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196157.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlCellType type</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type, object value)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "SpecialCells", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, type, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196157.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlCellType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "SpecialCells", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821663.aspx </remarks>
		/// <param name="edition">string edition</param>
		/// <param name="format">optional NetOffice.ExcelApi.Enums.XlSubscribeToFormat Format = -4158</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SubscribeTo(string edition, object format)
		{
			return Factory.ExecuteVariantMethodGet(this, "SubscribeTo", edition, format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821663.aspx </remarks>
		/// <param name="edition">string edition</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SubscribeTo(string edition)
		{
			return Factory.ExecuteVariantMethodGet(this, "SubscribeTo", edition);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
		/// <param name="groupBy">Int32 groupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
		/// <param name="totalList">object totalList</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="pageBreaks">optional object pageBreaks</param>
		/// <param name="summaryBelowData">optional NetOffice.ExcelApi.Enums.XlSummaryRow SummaryBelowData = 1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks, object summaryBelowData)
		{
			return Factory.ExecuteVariantMethodGet(this, "Subtotal", new object[]{ groupBy, function, totalList, replace, pageBreaks, summaryBelowData });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
		/// <param name="groupBy">Int32 groupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
		/// <param name="totalList">object totalList</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList)
		{
			return Factory.ExecuteVariantMethodGet(this, "Subtotal", groupBy, function, totalList);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
		/// <param name="groupBy">Int32 groupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
		/// <param name="totalList">object totalList</param>
		/// <param name="replace">optional object replace</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace)
		{
			return Factory.ExecuteVariantMethodGet(this, "Subtotal", groupBy, function, totalList, replace);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838166.aspx </remarks>
		/// <param name="groupBy">Int32 groupBy</param>
		/// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
		/// <param name="totalList">object totalList</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="pageBreaks">optional object pageBreaks</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks)
		{
			return Factory.ExecuteVariantMethodGet(this, "Subtotal", new object[]{ groupBy, function, totalList, replace, pageBreaks });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834729.aspx </remarks>
		/// <param name="rowInput">optional object rowInput</param>
		/// <param name="columnInput">optional object columnInput</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Table(object rowInput, object columnInput)
		{
			return Factory.ExecuteVariantMethodGet(this, "Table", rowInput, columnInput);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834729.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Table()
		{
			return Factory.ExecuteVariantMethodGet(this, "Table");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834729.aspx </remarks>
		/// <param name="rowInput">optional object rowInput</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Table(object rowInput)
		{
			return Factory.ExecuteVariantMethodGet(this, "Table", rowInput);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object trailingMinusNumbers</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator, trailingMinusNumbers });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns()
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", destination, dataType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", destination, dataType, textQualifier);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", destination, dataType, textQualifier, consecutiveDelimiter);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193593.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextToColumns", new object[]{ destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837754.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Ungroup()
		{
			return Factory.ExecuteVariantMethodGet(this, "Ungroup");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835301.aspx </remarks>
		/// <param name="text">optional object text</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comment AddComment(object text)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Comment>(this, "AddComment", NetOffice.ExcelApi.Comment.LateBindingApiWrapperType, text);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835301.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comment AddComment()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Comment>(this, "AddComment", NetOffice.ExcelApi.Comment.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823037.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ClearComments()
		{
			 Factory.ExecuteMethod(this, "ClearComments");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822349.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SetPhonetic()
		{
			 Factory.ExecuteMethod(this, "SetPhonetic");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut()
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197435.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			return Factory.ExecuteVariantMethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object skipBlanks</param>
		/// <param name="transpose">optional object transpose</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste, object operation, object skipBlanks, object transpose)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PasteSpecial", paste, operation, skipBlanks, transpose);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial()
		{
			return Factory.ExecuteVariantMethodGet(this, "_PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PasteSpecial", paste);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste, object operation)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PasteSpecial", paste, operation);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
		/// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
		/// <param name="skipBlanks">optional object skipBlanks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object _PasteSpecial(object paste, object operation, object skipBlanks)
		{
			return Factory.ExecuteVariantMethodGet(this, "_PasteSpecial", paste, operation, skipBlanks);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838797.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Dirty()
		{
			 Factory.ExecuteMethod(this, "Dirty");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194108.aspx </remarks>
		/// <param name="speakDirection">optional object speakDirection</param>
		/// <param name="speakFormulas">optional object speakFormulas</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak(object speakDirection, object speakFormulas)
		{
			 Factory.ExecuteMethod(this, "Speak", speakDirection, speakFormulas);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194108.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak()
		{
			 Factory.ExecuteMethod(this, "Speak");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194108.aspx </remarks>
		/// <param name="speakDirection">optional object speakDirection</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Speak(object speakDirection)
		{
			 Factory.ExecuteMethod(this, "Speak", speakDirection);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut()
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			return Factory.ExecuteVariantMethodGet(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193823.aspx </remarks>
		/// <param name="columns">optional object columns</param>
		/// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void RemoveDuplicates(object columns, object header)
		{
			 Factory.ExecuteMethod(this, "RemoveDuplicates", columns, header);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193823.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void RemoveDuplicates()
		{
			 Factory.ExecuteMethod(this, "RemoveDuplicates");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193823.aspx </remarks>
		/// <param name="columns">optional object columns</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void RemoveDuplicates(object columns)
		{
			 Factory.ExecuteMethod(this, "RemoveDuplicates", columns);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="openAfterPublish">optional object openAfterPublish</param>
		/// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type, filename, quality);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type, filename, quality, includeDocProperties);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836441.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="openAfterPublish">optional object openAfterPublish</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835226.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object CalculateRowMajorOrder()
		{
			return Factory.ExecuteVariantMethodGet(this, "CalculateRowMajorOrder");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		/// <param name="color">optional object color</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle, object weight, object colorIndex, object color)
		{
			return Factory.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle, weight, colorIndex, color);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public object _BorderAround()
		{
			return Factory.ExecuteVariantMethodGet(this, "_BorderAround");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle)
		{
			return Factory.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle, object weight)
		{
			return Factory.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle, weight);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="lineStyle">optional object lineStyle</param>
		/// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
		/// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public object _BorderAround(object lineStyle, object weight, object colorIndex)
		{
			return Factory.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle, weight, colorIndex);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194741.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void ClearHyperlinks()
		{
			 Factory.ExecuteMethod(this, "ClearHyperlinks");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838963.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void AllocateChanges()
		{
			 Factory.ExecuteMethod(this, "AllocateChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837815.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void DiscardChanges()
		{
			 Factory.ExecuteMethod(this, "DiscardChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228715.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public void FlashFill()
		{
			 Factory.ExecuteMethod(this, "FlashFill");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Range>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Range>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Range>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Range>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.Range> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Range item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}