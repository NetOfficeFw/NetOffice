using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// IRange
    /// </summary>
    [SyntaxBypass]
    public class IRange_ : COMObject, NetOffice.ExcelApi.IRange_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IRange_() : base()
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", new object[] { rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
        {
            return get_Address(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute)
        {
            return get_Address(rowAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute)
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute, object referenceStyle)
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address", rowAbsolute, columnAbsolute, referenceStyle, external);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Address
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Address")]
        public virtual string Address(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressLocal", new object[] { rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        /// <param name="relativeTo">optional object relativeTo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        public virtual string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external, object relativeTo)
        {
            return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_AddressLocal(object rowAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        public virtual string AddressLocal(object rowAbsolute)
        {
            return get_AddressLocal(rowAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_AddressLocal(object rowAbsolute, object columnAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute, columnAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        public virtual string AddressLocal(object rowAbsolute, object columnAbsolute)
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute, columnAbsolute, referenceStyle);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        public virtual string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle)
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressLocal", rowAbsolute, columnAbsolute, referenceStyle, external);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_AddressLocal
        /// </summary>
        /// <param name="rowAbsolute">optional object rowAbsolute</param>
        /// <param name="columnAbsolute">optional object columnAbsolute</param>
        /// <param name="referenceStyle">optional NetOffice.ExcelApi.Enums.XlReferenceStyle referenceStyle</param>
        /// <param name="external">optional object external</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_AddressLocal")]
        public virtual string AddressLocal(object rowAbsolute, object columnAbsolute, object referenceStyle, object external)
        {
            return get_AddressLocal(rowAbsolute, columnAbsolute, referenceStyle, external);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Characters get_Characters(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", typeof(NetOffice.ExcelApi.Characters), start, length);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Characters")]
        public virtual NetOffice.ExcelApi.Characters Characters(object start, object length)
        {
            return get_Characters(start, length);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Characters get_Characters(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", typeof(NetOffice.ExcelApi.Characters), start);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Characters")]
        public virtual NetOffice.ExcelApi.Characters Characters(object start)
        {
            return get_Characters(start);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_Offset(object rowOffset, object columnOffset)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Offset", typeof(NetOffice.ExcelApi.Range), rowOffset, columnOffset);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Offset
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        /// <param name="columnOffset">optional object columnOffset</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Offset")]
        public virtual NetOffice.ExcelApi.Range Offset(object rowOffset, object columnOffset)
        {
            return get_Offset(rowOffset, columnOffset);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_Offset(object rowOffset)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Offset", typeof(NetOffice.ExcelApi.Range), rowOffset);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Offset
        /// </summary>
        /// <param name="rowOffset">optional object rowOffset</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Offset")]
        public virtual NetOffice.ExcelApi.Range Offset(object rowOffset)
        {
            return get_Offset(rowOffset);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowSize">optional object rowSize</param>
        /// <param name="columnSize">optional object columnSize</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_Resize(object rowSize, object columnSize)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Resize", typeof(NetOffice.ExcelApi.Range), rowSize, columnSize);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Resize
        /// </summary>
        /// <param name="rowSize">optional object rowSize</param>
        /// <param name="columnSize">optional object columnSize</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Resize")]
        public virtual NetOffice.ExcelApi.Range Resize(object rowSize, object columnSize)
        {
            return get_Resize(rowSize, columnSize);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="rowSize">optional object rowSize</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_Resize(object rowSize)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Resize", typeof(NetOffice.ExcelApi.Range), rowSize);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Resize
        /// </summary>
        /// <param name="rowSize">optional object rowSize</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Resize")]
        public virtual NetOffice.ExcelApi.Range Resize(object rowSize)
        {
            return get_Resize(rowSize);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Value(object rangeValueDataType)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value", rangeValueDataType);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Value(object rangeValueDataType, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Value", rangeValueDataType, value);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Alias for get_Value
        /// </summary>
        /// <param name="rangeValueDataType">optional object rangeValueDataType</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), Redirect("get_Value")]
        public virtual object Value(object rangeValueDataType)
        {
            return get_Value(rangeValueDataType);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// Interface IRange 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    public class IRange : NetOffice.ExcelApi.Behind.IRange_, NetOffice.ExcelApi.IRange
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.ExcelApi.IRange);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(IRange);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IRange() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AddIndent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "AddIndent");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "AddIndent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Address
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Address");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string AddressLocal
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AddressLocal");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Areas Areas
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Areas>(this, "Areas", typeof(NetOffice.ExcelApi.Areas));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Borders Borders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Borders>(this, "Borders", typeof(NetOffice.ExcelApi.Borders));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Cells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Cells", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Characters Characters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", typeof(NetOffice.ExcelApi.Characters));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Column
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Column");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range Columns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Columns", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ColumnWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ColumnWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ColumnWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range CurrentArray
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "CurrentArray", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range CurrentRegion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "CurrentRegion", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
        /// Custom Indexer
		/// </summary>
		/// <param name="rowIndex">object rowIndex</param>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public virtual object this[object rowIndex]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "_Default", rowIndex);
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "_Default", value, rowIndex);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="rowIndex">optional object rowIndex</param>
        /// <param name="columnIndex">optional object columnIndex</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual object this[object rowIndex, object columnIndex]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "_Default", rowIndex, columnIndex);
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "_Default", value, rowIndex, columnIndex);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Dependents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Dependents", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range DirectDependents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DirectDependents", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range DirectPrecedents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DirectPrecedents", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection direction</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_End(NetOffice.ExcelApi.Enums.XlDirection direction)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "End", typeof(NetOffice.ExcelApi.Range), direction);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_End
        /// </summary>
        /// <param name="direction">NetOffice.ExcelApi.Enums.XlDirection direction</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_End")]
        public virtual NetOffice.ExcelApi.Range End(NetOffice.ExcelApi.Enums.XlDirection direction)
        {
            return get_End(direction);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range EntireColumn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "EntireColumn", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range EntireRow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "EntireRow", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Font Font
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Font>(this, "Font", typeof(NetOffice.ExcelApi.Font));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Formula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Formula");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Formula", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FormulaArray
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaArray");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaArray", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlFormulaLabel FormulaLabel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlFormulaLabel>(this, "FormulaLabel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FormulaLabel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FormulaHidden
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaHidden");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaHidden", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FormulaLocal
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaLocal");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaLocal", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FormulaR1C1
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaR1C1");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaR1C1", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FormulaR1C1Local
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FormulaR1C1Local");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FormulaR1C1Local", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object HasArray
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasArray");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object HasFormula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasFormula");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Height
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Height");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Hidden
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Hidden");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Hidden", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object HorizontalAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HorizontalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object IndentLevel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IndentLevel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "IndentLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Interior Interior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Interior>(this, "Interior", typeof(NetOffice.ExcelApi.Interior));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Left
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Left");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ListHeaderRows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ListHeaderRows");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlLocationInTable LocationInTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlLocationInTable>(this, "LocationInTable");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Locked
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Locked");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Locked", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range MergeArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "MergeArea", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object MergeCells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "MergeCells");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "MergeCells", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Next
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Next", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object NumberFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "NumberFormat");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "NumberFormat", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object NumberFormatLocal
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "NumberFormatLocal");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "NumberFormatLocal", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Offset
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Offset", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Orientation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Orientation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Orientation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object OutlineLevel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "OutlineLevel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "OutlineLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PageBreak
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PageBreak");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageBreak", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotField PivotField
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "PivotField", typeof(NetOffice.ExcelApi.PivotField));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotItem PivotItem
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotItem>(this, "PivotItem", typeof(NetOffice.ExcelApi.PivotItem));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotTable PivotTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTable", typeof(NetOffice.ExcelApi.PivotTable));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Precedents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Precedents", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrefixCharacter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrefixCharacter");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Previous
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Previous", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.QueryTable QueryTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.QueryTable>(this, "QueryTable", typeof(NetOffice.ExcelApi.QueryTable));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_Range(object cell1, object cell2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", typeof(NetOffice.ExcelApi.Range), cell1, cell2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Range
        /// </summary>
        /// <param name="cell1">object cell1</param>
        /// <param name="cell2">optional object cell2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Range")]
        public virtual NetOffice.ExcelApi.Range Range(object cell1, object cell2)
        {
            return get_Range(cell1, cell2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range get_Range(object cell1)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", typeof(NetOffice.ExcelApi.Range), cell1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Range
        /// </summary>
        /// <param name="cell1">object cell1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Range")]
        public virtual NetOffice.ExcelApi.Range Range(object cell1)
        {
            return get_Range(cell1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Resize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Resize", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Row
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Row");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object RowHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RowHeight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RowHeight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Range Rows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Rows", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowDetail
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ShowDetail");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ShowDetail", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShrinkToFit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ShrinkToFit");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ShrinkToFit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SoundNote SoundNote
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SoundNote>(this, "SoundNote", typeof(NetOffice.ExcelApi.SoundNote));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Style
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Style");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Style", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Summary
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Summary");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Text");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Top
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Top");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object UseStandardHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UseStandardHeight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "UseStandardHeight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object UseStandardWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UseStandardWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "UseStandardWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Validation Validation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Validation>(this, "Validation", typeof(NetOffice.ExcelApi.Validation));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Value
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Value", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Value2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value2");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Value2", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object VerticalAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "VerticalAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "VerticalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Width
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Width");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Worksheet Worksheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Worksheet>(this, "Worksheet", typeof(NetOffice.ExcelApi.Worksheet));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object WrapText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "WrapText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "WrapText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Comment Comment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Comment>(this, "Comment", typeof(NetOffice.ExcelApi.Comment));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Phonetic Phonetic
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Phonetic>(this, "Phonetic", typeof(NetOffice.ExcelApi.Phonetic));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.FormatConditions FormatConditions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.FormatConditions>(this, "FormatConditions", typeof(NetOffice.ExcelApi.FormatConditions));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ReadingOrder
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ReadingOrder");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadingOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Hyperlinks Hyperlinks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Hyperlinks>(this, "Hyperlinks", typeof(NetOffice.ExcelApi.Hyperlinks));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Phonetics Phonetics
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Phonetics>(this, "Phonetics", typeof(NetOffice.ExcelApi.Phonetics));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string ID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ID");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ID", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotCell PivotCell
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotCell>(this, "PivotCell", typeof(NetOffice.ExcelApi.PivotCell));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Errors Errors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Errors>(this, "Errors", typeof(NetOffice.ExcelApi.Errors));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SmartTags SmartTags
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SmartTags>(this, "SmartTags", typeof(NetOffice.ExcelApi.SmartTags));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool AllowEdit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowEdit");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.ListObject ListObject
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObject>(this, "ListObject", typeof(NetOffice.ExcelApi.ListObject));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.XPath XPath
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XPath>(this, "XPath", typeof(NetOffice.ExcelApi.XPath));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Actions ServerActions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Actions>(this, "ServerActions", typeof(NetOffice.ExcelApi.Actions));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual string MDX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MDX");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object CountLarge
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CountLarge");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SparklineGroups SparklineGroups
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SparklineGroups>(this, "SparklineGroups", typeof(NetOffice.ExcelApi.SparklineGroups));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.DisplayFormat DisplayFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DisplayFormat>(this, "DisplayFormat", typeof(NetOffice.ExcelApi.DisplayFormat));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Activate()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        /// <param name="criteriaRange">optional object criteriaRange</param>
        /// <param name="copyToRange">optional object copyToRange</param>
        /// <param name="unique">optional object unique</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange, object unique)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AdvancedFilter", action, criteriaRange, copyToRange, unique);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AdvancedFilter", action);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        /// <param name="criteriaRange">optional object criteriaRange</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AdvancedFilter", action, criteriaRange);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="action">NetOffice.ExcelApi.Enums.XlFilterAction action</param>
        /// <param name="criteriaRange">optional object criteriaRange</param>
        /// <param name="copyToRange">optional object copyToRange</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AdvancedFilter(NetOffice.ExcelApi.Enums.XlFilterAction action, object criteriaRange, object copyToRange)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AdvancedFilter", action, criteriaRange, copyToRange);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        /// <param name="omitRow">optional object omitRow</param>
        /// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
        /// <param name="appendLast">optional object appendLast</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order, object appendLast)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", new object[] { names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order, appendLast });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", names);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names, object ignoreRelativeAbsolute)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", names, ignoreRelativeAbsolute);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", names, ignoreRelativeAbsolute, useRowColumnNames);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        /// <param name="omitRow">optional object omitRow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", new object[] { names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="names">optional object names</param>
        /// <param name="ignoreRelativeAbsolute">optional object ignoreRelativeAbsolute</param>
        /// <param name="useRowColumnNames">optional object useRowColumnNames</param>
        /// <param name="omitColumn">optional object omitColumn</param>
        /// <param name="omitRow">optional object omitRow</param>
        /// <param name="order">optional NetOffice.ExcelApi.Enums.XlApplyNamesOrder Order = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyNames(object names, object ignoreRelativeAbsolute, object useRowColumnNames, object omitColumn, object omitRow, object order)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyNames", new object[] { names, ignoreRelativeAbsolute, useRowColumnNames, omitColumn, omitRow, order });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ApplyOutlineStyles()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyOutlineStyles");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="_string">string string</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string AutoComplete(string _string)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "AutoComplete", _string);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">NetOffice.ExcelApi.Range destination</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlAutoFillType Type = 0</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFill(NetOffice.ExcelApi.Range destination, object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFill", destination, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">NetOffice.ExcelApi.Range destination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFill(NetOffice.ExcelApi.Range destination)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFill", destination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
        /// <param name="criteria2">optional object criteria2</param>
        /// <param name="visibleDropDown">optional object visibleDropDown</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFilter(object field, object criteria1, object _operator, object criteria2, object visibleDropDown)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFilter", new object[] { field, criteria1, _operator, criteria2, visibleDropDown });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFilter()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFilter");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">optional object field</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFilter(object field)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFilter", field);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFilter(object field, object criteria1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFilter", field, criteria1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFilter(object field, object criteria1, object _operator)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFilter", field, criteria1, _operator);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="field">optional object field</param>
        /// <param name="criteria1">optional object criteria1</param>
        /// <param name="_operator">optional NetOffice.ExcelApi.Enums.XlAutoFilterOperator Operator = 1</param>
        /// <param name="criteria2">optional object criteria2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFilter(object field, object criteria1, object _operator, object criteria2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFilter", field, criteria1, _operator, criteria2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFit()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFit");
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format, object number, object font, object alignment, object border, object pattern, object width)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", new object[] { format, number, font, alignment, border, pattern, width });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", format);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format, object number)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", format, number);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format, object number, object font)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", format, number, font);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlRangeAutoFormat Format = 1</param>
        /// <param name="number">optional object number</param>
        /// <param name="font">optional object font</param>
        /// <param name="alignment">optional object alignment</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format, object number, object font, object alignment)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", format, number, font, alignment);
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format, object number, object font, object alignment, object border)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", new object[] { format, number, font, alignment, border });
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoFormat(object format, object number, object font, object alignment, object border, object pattern)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoFormat", new object[] { format, number, font, alignment, border, pattern });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object AutoOutline()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AutoOutline");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object BorderAround(object lineStyle, object weight, object colorIndex, object color)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BorderAround", lineStyle, weight, colorIndex, color);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        /// <param name="themeColor">optional object themeColor</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual object BorderAround(object lineStyle, object weight, object colorIndex, object color, object themeColor)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BorderAround", new object[] { lineStyle, weight, colorIndex, color, themeColor });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object BorderAround()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BorderAround");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object BorderAround(object lineStyle)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BorderAround", lineStyle);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object BorderAround(object lineStyle, object weight)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BorderAround", lineStyle, weight);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object BorderAround(object lineStyle, object weight, object colorIndex)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BorderAround", lineStyle, weight, colorIndex);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Calculate()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Calculate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="spellLang">optional object spellLang</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckSpelling()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckSpelling");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckSpelling(object customDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckSpelling(object customDictionary, object ignoreUppercase)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Clear()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Clear");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ClearContents()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ClearContents");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ClearFormats()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ClearFormats");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ClearNotes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ClearNotes");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ClearOutline()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ClearOutline");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="comparison">object comparison</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range ColumnDifferences(object comparison)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "ColumnDifferences", typeof(NetOffice.ExcelApi.Range), comparison);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        /// <param name="topRow">optional object topRow</param>
        /// <param name="leftColumn">optional object leftColumn</param>
        /// <param name="createLinks">optional object createLinks</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Consolidate(object sources, object function, object topRow, object leftColumn, object createLinks)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Consolidate", new object[] { sources, function, topRow, leftColumn, createLinks });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Consolidate()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Consolidate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sources">optional object sources</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Consolidate(object sources)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Consolidate", sources);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Consolidate(object sources, object function)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Consolidate", sources, function);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        /// <param name="topRow">optional object topRow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Consolidate(object sources, object function, object topRow)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Consolidate", sources, function, topRow);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sources">optional object sources</param>
        /// <param name="function">optional object function</param>
        /// <param name="topRow">optional object topRow</param>
        /// <param name="leftColumn">optional object leftColumn</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Consolidate(object sources, object function, object topRow, object leftColumn)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Consolidate", sources, function, topRow, leftColumn);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Copy(object destination)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy", destination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Copy()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        /// <param name="maxColumns">optional object maxColumns</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CopyFromRecordset(object data, object maxRows, object maxColumns)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows, maxColumns);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="data">object data</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CopyFromRecordset(object data)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CopyFromRecordset", data);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="data">object data</param>
        /// <param name="maxRows">optional object maxRows</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CopyFromRecordset(object data, object maxRows)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CopyFromRecordset", data, maxRows);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlCopyPictureFormat Format = -4147</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CopyPicture(object appearance, object format)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CopyPicture", appearance, format);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CopyPicture()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CopyPicture");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CopyPicture(object appearance)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CopyPicture", appearance);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        /// <param name="bottom">optional object bottom</param>
        /// <param name="right">optional object right</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreateNames(object top, object left, object bottom, object right)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreateNames", top, left, bottom, right);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreateNames()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreateNames");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="top">optional object top</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreateNames(object top)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreateNames", top);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreateNames(object top, object left)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreateNames", top, left);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        /// <param name="bottom">optional object bottom</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreateNames(object top, object left, object bottom)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreateNames", top, left, bottom);
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF, object containsVALU)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher", new object[] { edition, appearance, containsPICT, containsBIFF, containsRTF, containsVALU });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher(object edition)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher", edition);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher(object edition, object appearance)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher", edition, appearance);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher(object edition, object appearance, object containsPICT)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher", edition, appearance, containsPICT);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsBIFF">optional object containsBIFF</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher", edition, appearance, containsPICT, containsBIFF);
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CreatePublisher(object edition, object appearance, object containsPICT, object containsBIFF, object containsRTF)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CreatePublisher", new object[] { edition, appearance, containsPICT, containsBIFF, containsRTF });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Cut(object destination)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Cut", destination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Cut()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        /// <param name="step">optional object step</param>
        /// <param name="stop">optional object stop</param>
        /// <param name="trend">optional object trend</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries(object rowcol, object type, object date, object step, object stop, object trend)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries", new object[] { rowcol, type, date, step, stop, trend });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional object rowcol</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries(object rowcol)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries", rowcol);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries(object rowcol, object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries", rowcol, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries(object rowcol, object type, object date)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries", rowcol, type, date);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        /// <param name="step">optional object step</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries(object rowcol, object type, object date, object step)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries", rowcol, type, date, step);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlDataSeriesType Type = -4132</param>
        /// <param name="date">optional NetOffice.ExcelApi.Enums.XlDataSeriesDate Date = 1</param>
        /// <param name="step">optional object step</param>
        /// <param name="stop">optional object stop</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DataSeries(object rowcol, object type, object date, object step, object stop)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataSeries", new object[] { rowcol, type, date, step, stop });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Delete(object shift)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete", shift);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Delete()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DialogBox()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DialogBox");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
        /// <param name="format">optional object format</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize, object format)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditionOptions", new object[] { type, option, name, reference, appearance, chartSize, format });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditionOptions", type, option);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditionOptions", type, option, name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditionOptions", type, option, name, reference);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditionOptions", new object[] { type, option, name, reference, appearance });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlEditionType type</param>
        /// <param name="option">NetOffice.ExcelApi.Enums.XlEditionOptionsOption option</param>
        /// <param name="name">optional object name</param>
        /// <param name="reference">optional object reference</param>
        /// <param name="appearance">optional NetOffice.ExcelApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="chartSize">optional NetOffice.ExcelApi.Enums.XlPictureAppearance ChartSize = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditionOptions(NetOffice.ExcelApi.Enums.XlEditionType type, NetOffice.ExcelApi.Enums.XlEditionOptionsOption option, object name, object reference, object appearance, object chartSize)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditionOptions", new object[] { type, option, name, reference, appearance, chartSize });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FillDown()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FillDown");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FillLeft()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FillLeft");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FillRight()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FillRight");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FillUp()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FillUp");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        /// <param name="searchFormat">optional object searchFormat</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase, object matchByte, object searchFormat)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, searchFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), what);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), what, after);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), what, after, lookIn);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), what, after, lookIn, lookAt);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), new object[] { what, after, lookIn, lookAt, searchOrder });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="after">optional object after</param>
        /// <param name="lookIn">optional object lookIn</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="searchDirection">optional NetOffice.ExcelApi.Enums.XlSearchDirection SearchDirection = 1</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range Find(object what, object after, object lookIn, object lookAt, object searchOrder, object searchDirection, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "Find", typeof(NetOffice.ExcelApi.Range), new object[] { what, after, lookIn, lookAt, searchOrder, searchDirection, matchCase });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range FindNext(object after)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindNext", typeof(NetOffice.ExcelApi.Range), after);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range FindNext()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindNext", typeof(NetOffice.ExcelApi.Range));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range FindPrevious(object after)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindPrevious", typeof(NetOffice.ExcelApi.Range), after);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range FindPrevious()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "FindPrevious", typeof(NetOffice.ExcelApi.Range));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object FunctionWizard()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FunctionWizard");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="goal">object goal</param>
        /// <param name="changingCell">NetOffice.ExcelApi.Range changingCell</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool GoalSeek(object goal, NetOffice.ExcelApi.Range changingCell)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GoalSeek", goal, changingCell);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="end">optional object end</param>
        /// <param name="by">optional object by</param>
        /// <param name="periods">optional object periods</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Group(object start, object end, object by, object periods)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Group", start, end, by, periods);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Group()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Group");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="start">optional object start</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Group(object start)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Group", start);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="end">optional object end</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Group(object start, object end)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Group", start, end);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="end">optional object end</param>
        /// <param name="by">optional object by</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Group(object start, object end, object by)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Group", start, end, by);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="insertAmount">Int32 insertAmount</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 InsertIndent(Int32 insertAmount)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InsertIndent", insertAmount);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="shift">optional object shift</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Insert(object shift)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Insert", shift);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="shift">optional object shift</param>
        /// <param name="copyOrigin">optional object copyOrigin</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Insert(object shift, object copyOrigin)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Insert", shift, copyOrigin);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Insert()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Insert");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Justify()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Justify");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ListNames()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ListNames");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="across">optional object across</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Merge(object across)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Merge", across);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Merge()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Merge");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 UnMerge()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UnMerge");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="towardPrecedent">optional object towardPrecedent</param>
        /// <param name="arrowNumber">optional object arrowNumber</param>
        /// <param name="linkNumber">optional object linkNumber</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object NavigateArrow(object towardPrecedent, object arrowNumber, object linkNumber)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "NavigateArrow", towardPrecedent, arrowNumber, linkNumber);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object NavigateArrow()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "NavigateArrow");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="towardPrecedent">optional object towardPrecedent</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object NavigateArrow(object towardPrecedent)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "NavigateArrow", towardPrecedent);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="towardPrecedent">optional object towardPrecedent</param>
        /// <param name="arrowNumber">optional object arrowNumber</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object NavigateArrow(object towardPrecedent, object arrowNumber)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "NavigateArrow", towardPrecedent, arrowNumber);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">optional object text</param>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string NoteText(object text, object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "NoteText", text, start, length);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string NoteText()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "NoteText");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">optional object text</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string NoteText(object text)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "NoteText", text);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">optional object text</param>
        /// <param name="start">optional object start</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string NoteText(object text, object start)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "NoteText", text, start);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="parseLine">optional object parseLine</param>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Parse(object parseLine, object destination)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Parse", parseLine, destination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Parse()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Parse");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="parseLine">optional object parseLine</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Parse(object parseLine)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Parse", parseLine);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        /// <param name="transpose">optional object transpose</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PasteSpecial(object paste, object operation, object skipBlanks, object transpose)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PasteSpecial", paste, operation, skipBlanks, transpose);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PasteSpecial()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PasteSpecial");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PasteSpecial(object paste)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PasteSpecial", paste);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PasteSpecial(object paste, object operation)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PasteSpecial", paste, operation);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PasteSpecial(object paste, object operation, object skipBlanks)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PasteSpecial", paste, operation, skipBlanks);
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to, object copies)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", from, to, copies);
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to, object copies, object preview)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", from, to, copies, preview);
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter });
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="enableChanges">optional object enableChanges</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintPreview(object enableChanges)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintPreview", enableChanges);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintPreview()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintPreview");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object RemoveSubtotal()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RemoveSubtotal");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", new object[] { what, replacement, lookAt, searchOrder, matchCase, matchByte });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        /// <param name="searchFormat">optional object searchFormat</param>
        /// <param name="replaceFormat">optional object replaceFormat</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat, object replaceFormat)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", new object[] { what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat, replaceFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", what, replacement);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement, object lookAt)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", what, replacement, lookAt);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement, object lookAt, object searchOrder)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", what, replacement, lookAt, searchOrder);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", new object[] { what, replacement, lookAt, searchOrder, matchCase });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="what">object what</param>
        /// <param name="replacement">object replacement</param>
        /// <param name="lookAt">optional object lookAt</param>
        /// <param name="searchOrder">optional object searchOrder</param>
        /// <param name="matchCase">optional object matchCase</param>
        /// <param name="matchByte">optional object matchByte</param>
        /// <param name="searchFormat">optional object searchFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool Replace(object what, object replacement, object lookAt, object searchOrder, object matchCase, object matchByte, object searchFormat)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Replace", new object[] { what, replacement, lookAt, searchOrder, matchCase, matchByte, searchFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="comparison">object comparison</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range RowDifferences(object comparison)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "RowDifferences", typeof(NetOffice.ExcelApi.Range), comparison);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Select()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Show()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Show");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="remove">optional object remove</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowDependents(object remove)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowDependents", remove);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowDependents()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowDependents");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowErrors()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowErrors");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="remove">optional object remove</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowPrecedents(object remove)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowPrecedents", remove);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ShowPrecedents()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ShowPrecedents");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2, object dataOption3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2, dataOption3 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", key1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", key1, order1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", key1, order1, key2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", key1, order1, key2, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="type">optional object type</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object Sort(object key1, object order1, object key2, object type, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object sortMethod, object dataOption1, object dataOption2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Sort", new object[] { key1, order1, key2, type, order2, key3, order3, header, orderCustom, matchCase, orientation, sortMethod, dataOption1, dataOption2 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2, object dataOption3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2, dataOption3 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod, key1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod, key1, order1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", sortMethod, key1, order1, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortMethod">optional NetOffice.ExcelApi.Enums.XlSortMethod SortMethod = 1</param>
        /// <param name="key1">optional object key1</param>
        /// <param name="order1">optional NetOffice.ExcelApi.Enums.XlSortOrder Order1 = 1</param>
        /// <param name="type">optional object type</param>
        /// <param name="key2">optional object key2</param>
        /// <param name="order2">optional NetOffice.ExcelApi.Enums.XlSortOrder Order2 = 1</param>
        /// <param name="key3">optional object key3</param>
        /// <param name="order3">optional NetOffice.ExcelApi.Enums.XlSortOrder Order3 = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1 });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object SortSpecial(object sortMethod, object key1, object order1, object type, object key2, object order2, object key3, object order3, object header, object orderCustom, object matchCase, object orientation, object dataOption1, object dataOption2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SortSpecial", new object[] { sortMethod, key1, order1, type, key2, order2, key3, order3, header, orderCustom, matchCase, orientation, dataOption1, dataOption2 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlCellType type</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type, object value)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "SpecialCells", typeof(NetOffice.ExcelApi.Range), type, value);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlCellType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Range SpecialCells(NetOffice.ExcelApi.Enums.XlCellType type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "SpecialCells", typeof(NetOffice.ExcelApi.Range), type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">string edition</param>
        /// <param name="format">optional NetOffice.ExcelApi.Enums.XlSubscribeToFormat Format = -4158</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SubscribeTo(string edition, object format)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SubscribeTo", edition, format);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">string edition</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SubscribeTo(string edition)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SubscribeTo", edition);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        /// <param name="replace">optional object replace</param>
        /// <param name="pageBreaks">optional object pageBreaks</param>
        /// <param name="summaryBelowData">optional NetOffice.ExcelApi.Enums.XlSummaryRow SummaryBelowData = 1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks, object summaryBelowData)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Subtotal", new object[] { groupBy, function, totalList, replace, pageBreaks, summaryBelowData });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Subtotal", groupBy, function, totalList);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        /// <param name="replace">optional object replace</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Subtotal", groupBy, function, totalList, replace);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="groupBy">Int32 groupBy</param>
        /// <param name="function">NetOffice.ExcelApi.Enums.XlConsolidationFunction function</param>
        /// <param name="totalList">object totalList</param>
        /// <param name="replace">optional object replace</param>
        /// <param name="pageBreaks">optional object pageBreaks</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Subtotal(Int32 groupBy, NetOffice.ExcelApi.Enums.XlConsolidationFunction function, object totalList, object replace, object pageBreaks)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Subtotal", new object[] { groupBy, function, totalList, replace, pageBreaks });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowInput">optional object rowInput</param>
        /// <param name="columnInput">optional object columnInput</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Table(object rowInput, object columnInput)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Table", rowInput, columnInput);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Table()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Table");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowInput">optional object rowInput</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Table(object rowInput)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Table", rowInput);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator, thousandsSeparator, trailingMinusNumbers });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", destination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", destination, dataType);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", destination, dataType, textQualifier);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", destination, dataType, textQualifier, consecutiveDelimiter);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="dataType">optional NetOffice.ExcelApi.Enums.XlTextParsingType DataType = 1</param>
        /// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
        /// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
        /// <param name="tab">optional object tab</param>
        /// <param name="semicolon">optional object semicolon</param>
        /// <param name="comma">optional object comma</param>
        /// <param name="space">optional object space</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextToColumns(object destination, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object decimalSeparator)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextToColumns", new object[] { destination, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, decimalSeparator });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Ungroup()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Ungroup");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">optional object text</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Comment AddComment(object text)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Comment>(this, "AddComment", typeof(NetOffice.ExcelApi.Comment), text);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Comment AddComment()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Comment>(this, "AddComment", typeof(NetOffice.ExcelApi.Comment));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ClearComments()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearComments");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SetPhonetic()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetPhonetic");
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
        /// <param name="prToFileName">optional object prToFileName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to, object copies)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", from, to, copies);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to, object copies, object preview)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", from, to, copies, preview);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter });
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
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
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
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        /// <param name="transpose">optional object transpose</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object _PasteSpecial(object paste, object operation, object skipBlanks, object transpose)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PasteSpecial", paste, operation, skipBlanks, transpose);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object _PasteSpecial()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PasteSpecial");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object _PasteSpecial(object paste)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PasteSpecial", paste);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object _PasteSpecial(object paste, object operation)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PasteSpecial", paste, operation);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="paste">optional NetOffice.ExcelApi.Enums.XlPasteType Paste = -4104</param>
        /// <param name="operation">optional NetOffice.ExcelApi.Enums.XlPasteSpecialOperation Operation = -4142</param>
        /// <param name="skipBlanks">optional object skipBlanks</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual object _PasteSpecial(object paste, object operation, object skipBlanks)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_PasteSpecial", paste, operation, skipBlanks);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Dirty()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Dirty");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="speakDirection">optional object speakDirection</param>
        /// <param name="speakFormulas">optional object speakFormulas</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Speak(object speakDirection, object speakFormulas)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak", speakDirection, speakFormulas);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Speak()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="speakDirection">optional object speakDirection</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Speak(object speakDirection)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak", speakDirection);
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from, object to, object copies)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", from, to, copies);
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from, object to, object copies, object preview)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", from, to, copies, preview);
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter });
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="columns">optional object columns</param>
        /// <param name="header">optional NetOffice.ExcelApi.Enums.XlYesNoGuess Header = 2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 RemoveDuplicates(object columns, object header)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveDuplicates", columns, header);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 RemoveDuplicates()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveDuplicates");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="columns">optional object columns</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 RemoveDuplicates(object columns)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveDuplicates", columns);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        /// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename, quality);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename, quality, includeDocProperties);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object CalculateRowMajorOrder()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CalculateRowMajorOrder");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        /// <param name="color">optional object color</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual object _BorderAround(object lineStyle, object weight, object colorIndex, object color)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle, weight, colorIndex, color);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual object _BorderAround()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_BorderAround");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual object _BorderAround(object lineStyle)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual object _BorderAround(object lineStyle, object weight)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle, weight);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="lineStyle">optional object lineStyle</param>
        /// <param name="weight">optional NetOffice.ExcelApi.Enums.XlBorderWeight Weight = 2</param>
        /// <param name="colorIndex">optional NetOffice.ExcelApi.Enums.XlColorIndex ColorIndex = -4105</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual object _BorderAround(object lineStyle, object weight, object colorIndex)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_BorderAround", lineStyle, weight, colorIndex);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 ClearHyperlinks()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearHyperlinks");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 AllocateChanges()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AllocateChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 DiscardChanges()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DiscardChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public virtual Int32 FlashFill()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FlashFill");
        }

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, true);
        }

        #endregion

        #pragma warning restore
    }
}

