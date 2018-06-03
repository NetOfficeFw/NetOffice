using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IControlFormat 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IControlFormat : COMObject, NetOffice.ExcelApi.IControlFormat
    {
        #pragma warning disable

        #region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IControlFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IControlFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 DropDownLines
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "DropDownLines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DropDownLines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool Enabled
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Enabled");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Enabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 LargeChange
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "LargeChange");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LargeChange", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string LinkedCell
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "LinkedCell");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LinkedCell", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 ListCount
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "ListCount");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ListCount", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string ListFillRange
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ListFillRange");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ListFillRange", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 ListIndex
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "ListIndex");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ListIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool LockedText
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "LockedText");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LockedText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Max
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Max");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Max", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Min
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Min");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Min", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 MultiSelect
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "MultiSelect");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MultiSelect", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool PrintObject
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "PrintObject");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PrintObject", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 SmallChange
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "SmallChange");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SmallChange", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 _Default
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "_Default");
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Value
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Value");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Value", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">string text</param>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddItem(string text, object index)
        {
            return Factory.ExecuteInt32MethodGet(this, "AddItem", text, index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="text">string text</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddItem(string text)
        {
            return Factory.ExecuteInt32MethodGet(this, "AddItem", text);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 RemoveAllItems()
        {
            return Factory.ExecuteInt32MethodGet(this, "RemoveAllItems");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 RemoveItem(Int32 index, object count)
        {
            return Factory.ExecuteInt32MethodGet(this, "RemoveItem", index, count);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">Int32 index</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 RemoveItem(Int32 index)
        {
            return Factory.ExecuteInt32MethodGet(this, "RemoveItem", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object List(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "List", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object List()
        {
            return Factory.ExecuteVariantMethodGet(this, "List");
        }

        #endregion

        #pragma warning restore
    }
}

