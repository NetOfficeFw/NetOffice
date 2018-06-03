using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IDummy 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IDummy : COMObject, NetOffice.ExcelApi.IDummy
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
                    _type = typeof(IDummy);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDummy() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool ShowSignaturesPane
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowSignaturesPane");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowSignaturesPane", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 _ActiveSheetOrChart()
        {
            return Factory.ExecuteInt32MethodGet(this, "_ActiveSheetOrChart");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 RGB()
        {
            return Factory.ExecuteInt32MethodGet(this, "RGB");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 ChDir()
        {
            return Factory.ExecuteInt32MethodGet(this, "ChDir");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 DoScript()
        {
            return Factory.ExecuteInt32MethodGet(this, "DoScript");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 DirectObject()
        {
            return Factory.ExecuteInt32MethodGet(this, "DirectObject");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 RefreshDocument()
        {
            return Factory.ExecuteInt32MethodGet(this, "RefreshDocument");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="sigProv">object sigProv</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Signature AddSignatureLine(object sigProv)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "AddSignatureLine", typeof(NetOffice.OfficeApi.Signature), sigProv);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="sigProv">object sigProv</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Signature AddNonVisibleSignature(object sigProv)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "AddNonVisibleSignature", typeof(NetOffice.OfficeApi.Signature), sigProv);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 ThemeFontScheme()
        {
            return Factory.ExecuteInt32MethodGet(this, "ThemeFontScheme");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 ThemeColorScheme()
        {
            return Factory.ExecuteInt32MethodGet(this, "ThemeColorScheme");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 ThemeEffectScheme()
        {
            return Factory.ExecuteInt32MethodGet(this, "ThemeEffectScheme");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Int32 Load()
        {
            return Factory.ExecuteInt32MethodGet(this, "Load");
        }

        #endregion

        #pragma warning restore
    }
}
