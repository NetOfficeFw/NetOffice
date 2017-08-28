using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// Interface IConverterUICallback 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863370.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IConverterUICallback : COMObject
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
                    _type = typeof(IConverterUICallback);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IConverterUICallback(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IConverterUICallback(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterUICallback(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterUICallback(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterUICallback(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterUICallback(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterUICallback() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IConverterUICallback(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861826.aspx </remarks>
		/// <param name="uPercentComplete">UIntPtr uPercentComplete</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrReportProgress(UIntPtr uPercentComplete)
		{
			return Factory.ExecuteInt32MethodGet(this, "HrReportProgress", uPercentComplete);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861376.aspx </remarks>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="uType">UIntPtr uType</param>
		/// <param name="pidResult">Int32 pidResult</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrMessageBox(string bstrText, string bstrCaption, UIntPtr uType, out Int32 pidResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			pidResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrText, bstrCaption, uType, pidResult);
			object returnItem = Invoker.MethodReturn(this, "HrMessageBox", paramsArray, modifiers);
			pidResult = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861803.aspx </remarks>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="pbstrInput">string pbstrInput</param>
		/// <param name="fPassword">Int32 fPassword</param>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 HrInputBox(string bstrText, string bstrCaption, out string pbstrInput, Int32 fPassword)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,false);
			pbstrInput = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(bstrText, bstrCaption, pbstrInput, fPassword);
			object returnItem = Invoker.MethodReturn(this, "HrInputBox", paramsArray, modifiers);
			pbstrInput = paramsArray[2] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
