using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// Interface IConverter 
	/// SupportByVersion Office, 14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861235.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IConverter : COMObject
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
                    _type = typeof(IConverter);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864088.aspx
		/// </summary>
		/// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
		/// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
		/// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public Int32 HrInitConverter(NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrInitConverter", paramsArray);
			ppcp = (NetOffice.OfficeApi.IConverterPreferences)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862058.aspx
		/// </summary>
		/// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public Int32 HrUninitConverter(NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrUninitConverter", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864636.aspx
		/// </summary>
		/// <param name="bstrSourcePath">string bstrSourcePath</param>
		/// <param name="bstrDestPath">string bstrDestPath</param>
		/// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
		/// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
		/// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public Int32 HrImport(string bstrSourcePath, string bstrDestPath, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,false);
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSourcePath, bstrDestPath, pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrImport", paramsArray);
			ppcp = (NetOffice.OfficeApi.IConverterPreferences)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863696.aspx
		/// </summary>
		/// <param name="bstrSourcePath">string bstrSourcePath</param>
		/// <param name="bstrDestPath">string bstrDestPath</param>
		/// <param name="bstrClass">string bstrClass</param>
		/// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
		/// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
		/// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public Int32 HrExport(string bstrSourcePath, string bstrDestPath, string bstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false);
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSourcePath, bstrDestPath, bstrClass, pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrExport", paramsArray);
			ppcp = (NetOffice.OfficeApi.IConverterPreferences)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864094.aspx
		/// </summary>
		/// <param name="bstrPath">string bstrPath</param>
		/// <param name="pbstrClass">string pbstrClass</param>
		/// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
		/// <param name="ppcp">NetOffice.OfficeApi.IConverterPreferences ppcp</param>
		/// <param name="pcuic">NetOffice.OfficeApi.IConverterUICallback pcuic</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public Int32 HrGetFormat(string bstrPath, out string pbstrClass, NetOffice.OfficeApi.IConverterApplicationPreferences pcap, out NetOffice.OfficeApi.IConverterPreferences ppcp, NetOffice.OfficeApi.IConverterUICallback pcuic)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,true,false);
			pbstrClass = string.Empty;
			ppcp = null;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrPath, pbstrClass, pcap, ppcp, pcuic);
			object returnItem = Invoker.MethodReturn(this, "HrGetFormat", paramsArray);
			pbstrClass = (string)paramsArray[1];
			ppcp = (NetOffice.OfficeApi.IConverterPreferences)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861511.aspx
		/// </summary>
		/// <param name="hrErr">Int32 hrErr</param>
		/// <param name="pbstrErrorMsg">string pbstrErrorMsg</param>
		/// <param name="pcap">NetOffice.OfficeApi.IConverterApplicationPreferences pcap</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public Int32 HrGetErrorString(Int32 hrErr, out string pbstrErrorMsg, NetOffice.OfficeApi.IConverterApplicationPreferences pcap)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			pbstrErrorMsg = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(hrErr, pbstrErrorMsg, pcap);
			object returnItem = Invoker.MethodReturn(this, "HrGetErrorString", paramsArray);
			pbstrErrorMsg = (string)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}