using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface IConverter 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861235.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IConverter : COMObject, NetOffice.OfficeApi.IConverter
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
                    _contractType = typeof(NetOffice.OfficeApi.IConverter);
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
                    _type = typeof(IConverter);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IConverter() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864088.aspx </remarks>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrInitConverter(NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false);
            ppcp = null;
            object[] paramsArray = new object[] { pcap, ppcp, pcuic };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrInitConverter", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                ppcp = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.IConverterPreferences>(this, paramsArray[1], typeof(NetOffice.OfficeApi.IConverterPreferences));
            else
                ppcp = null;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862058.aspx </remarks>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrUninitConverter(NetOffice.OfficeApi.IConverterUICallback pcuic)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HrUninitConverter", pcuic);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864636.aspx </remarks>
        /// <param name="bstrSourcePath">string bstrSourcePath</param>
        /// <param name="bstrDestPath">string bstrDestPath</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrImport(string bstrSourcePath, string bstrDestPath, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, true, false);
            ppcp = null;
            object[] paramsArray = new object[]{ bstrSourcePath, bstrDestPath, pcap, ppcp, pcuic};

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrImport", paramsArray, modifiers);

            if (paramsArray[3] is MarshalByRefObject)
                ppcp = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.IConverterPreferences>(this, paramsArray[3], typeof(NetOffice.OfficeApi.IConverterPreferences));
            else
                ppcp = null;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863696.aspx </remarks>
        /// <param name="bstrSourcePath">string bstrSourcePath</param>
        /// <param name="bstrDestPath">string bstrDestPath</param>
        /// <param name="bstrClass">string bstrClass</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrExport(string bstrSourcePath, string bstrDestPath, string bstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, false, true, false);
            ppcp = null;
            object[] paramsArray = new object[] { bstrSourcePath, bstrDestPath, bstrClass, pcap, ppcp, pcuic };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrExport", paramsArray, modifiers);

            if (paramsArray[4] is MarshalByRefObject)
                ppcp = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.IConverterPreferences>(this, paramsArray[4], typeof(NetOffice.OfficeApi.IConverterPreferences));
            else
                ppcp = null;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864094.aspx </remarks>
        /// <param name="bstrPath">string bstrPath</param>
        /// <param name="pbstrClass">string pbstrClass</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        /// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
        /// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrGetFormat(string bstrPath, out string pbstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false, true, false);
            pbstrClass = string.Empty;
            ppcp = null;
            object[] paramsArray = new object[] { bstrPath, pbstrClass, pcap, ppcp, pcuic };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrGetFormat", paramsArray, modifiers);

            pbstrClass = (string)paramsArray[1];
            if (paramsArray[3] is MarshalByRefObject)
                ppcp = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.IConverterPreferences>(this, paramsArray[3], typeof(NetOffice.OfficeApi.IConverterPreferences));
            else
                ppcp = null;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861511.aspx </remarks>
        /// <param name="hrErr">Int32 hrErr</param>
        /// <param name="pbstrErrorMsg">string pbstrErrorMsg</param>
        /// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrGetErrorString(Int32 hrErr, out string pbstrErrorMsg, NetOffice.OfficeApi.IConverterApplicationPreferences pcap)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false);
            pbstrErrorMsg = string.Empty;
            object[] paramsArray = new object[] { hrErr, pbstrErrorMsg, pcap };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrGetErrorString", paramsArray, modifiers);

            pbstrErrorMsg = paramsArray[1] as string;
            return returnItem;
        }

        #endregion

        #pragma warning restore
    }
}
