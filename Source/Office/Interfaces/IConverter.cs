using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// Interface IConverter 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861235.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IConverter : COMObject
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
                    _type = typeof(IConverter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IConverter(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IConverter(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverter(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverter(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverter(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverter(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverter() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverter(string progId) : base(progId)
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
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrInitConverter(NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrInitConverter", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppcp = new NetOffice.OfficeApi.IConverterPreferences(this, paramsArray[1]);
            else
                ppcp = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862058.aspx </remarks>
		/// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrUninitConverter(NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			return Factory.ExecuteInt32MethodGet(this, "HrUninitConverter", pcuic);
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
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrImport(string bstrSourcePath, string bstrDestPath, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,false);
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSourcePath, bstrDestPath, pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrImport", paramsArray, modifiers);
            if (paramsArray[3] is MarshalByRefObject)
                ppcp = new NetOffice.OfficeApi.IConverterPreferences(this, paramsArray[3]);
            else
                ppcp = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
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
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrExport(string bstrSourcePath, string bstrDestPath, string bstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false);
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSourcePath, bstrDestPath, bstrClass, pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrExport", paramsArray, modifiers);
            if (paramsArray[4] is MarshalByRefObject)
                ppcp = new NetOffice.OfficeApi.IConverterPreferences(this, paramsArray[4]);
            else
                ppcp = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
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
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrGetFormat(string bstrPath, out string pbstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,true,false);
			pbstrClass = string.Empty;
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrPath, pbstrClass, pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrGetFormat", paramsArray, modifiers);
			pbstrClass = (string)paramsArray[1];
            if (paramsArray[3] is MarshalByRefObject)
                ppcp = new NetOffice.OfficeApi.IConverterPreferences(this, paramsArray[3]);
            else
                ppcp = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861511.aspx </remarks>
		/// <param name="hrErr">Int32 hrErr</param>
		/// <param name="pbstrErrorMsg">string pbstrErrorMsg</param>
		/// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrGetErrorString(Int32 hrErr, out string pbstrErrorMsg, NetOffice.OfficeApi.IConverterApplicationPreferences pcap)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			pbstrErrorMsg = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(hrErr, pbstrErrorMsg, pcap);
			object returnItem = Invoker.MethodReturn(this, "HrGetErrorString", paramsArray, modifiers);
			pbstrErrorMsg = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
