using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IHyperlink 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IHyperlink : COMObject, NetOffice.ExcelApi.IHyperlink
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
                    _type = typeof(IHyperlink);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHyperlink() : base()
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
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Range Range
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Shape Shape
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shape>(this, "Shape", typeof(NetOffice.ExcelApi.Shape));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string SubAddress
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "SubAddress");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SubAddress", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Address
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Address");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Address", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Type
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string EmailSubject
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "EmailSubject");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "EmailSubject", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string ScreenTip
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ScreenTip");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ScreenTip", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string TextToDisplay
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "TextToDisplay");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TextToDisplay", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddToFavorites()
        {
            return Factory.ExecuteInt32MethodGet(this, "AddToFavorites");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Delete()
        {
            return Factory.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        /// <param name="headerInfo">optional object headerInfo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Follow(object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
        {
            return Factory.ExecuteInt32MethodGet(this, "Follow", new object[] { newWindow, addHistory, extraInfo, method, headerInfo });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Follow()
        {
            return Factory.ExecuteInt32MethodGet(this, "Follow");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Follow(object newWindow)
        {
            return Factory.ExecuteInt32MethodGet(this, "Follow", newWindow);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Follow(object newWindow, object addHistory)
        {
            return Factory.ExecuteInt32MethodGet(this, "Follow", newWindow, addHistory);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Follow(object newWindow, object addHistory, object extraInfo)
        {
            return Factory.ExecuteInt32MethodGet(this, "Follow", newWindow, addHistory, extraInfo);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Follow(object newWindow, object addHistory, object extraInfo, object method)
        {
            return Factory.ExecuteInt32MethodGet(this, "Follow", newWindow, addHistory, extraInfo, method);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="editNow">bool editNow</param>
        /// <param name="overwrite">bool overwrite</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 CreateNewDocument(string filename, bool editNow, bool overwrite)
        {
            return Factory.ExecuteInt32MethodGet(this, "CreateNewDocument", filename, editNow, overwrite);
        }

        #endregion

        #pragma warning restore
    }
}

