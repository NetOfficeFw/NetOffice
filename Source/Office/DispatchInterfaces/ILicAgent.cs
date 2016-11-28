using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface ILicAgent 
	/// SupportByVersion Office, 10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ILicAgent : COMObject
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
                    _type = typeof(ILicAgent);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ILicAgent(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicAgent(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwBPC">Int32 dwBPC</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="bstrLicSource">string bstrLicSource</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 Initialize(Int32 dwBPC, Int32 dwMode, string bstrLicSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwBPC, dwMode, bstrLicSource);
			object returnItem = Invoker.MethodReturn(this, "Initialize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetFirstName()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetFirstName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetFirstName(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetFirstName", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetLastName()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetLastName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetLastName(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetLastName", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetOrgName()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetOrgName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetOrgName(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetOrgName", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetEmail()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetEmail", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetEmail(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetEmail", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetPhone()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetPhone", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetPhone(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetPhone", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetAddress1()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAddress1", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetAddress1(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetAddress1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCity()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCity", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCity(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetCity", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetState()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetState", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetState(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetState", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCountryCode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCountryCode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCountryCode(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetCountryCode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCountryDesc()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCountryDesc", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCountryDesc(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetCountryDesc", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetZip()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetZip", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetZip(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetZip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetIsoLanguage()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetIsoLanguage", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwNewVal">Int32 dwNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetIsoLanguage(Int32 dwNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwNewVal);
			Invoker.Method(this, "SetIsoLanguage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetMSUpdate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetMSUpdate", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetMSUpdate(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetMSUpdate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetMSOffer()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetMSOffer", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetMSOffer(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetMSOffer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetOtherOffer()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetOtherOffer", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetOtherOffer(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetOtherOffer", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetAddress2()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAddress2", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetAddress2(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetAddress2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 CheckSystemClock()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CheckSystemClock", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public DateTime GetExistingExpiryDate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetExistingExpiryDate", paramsArray);
			return NetRuntimeSystem.Convert.ToDateTime(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public DateTime GetNewExpiryDate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetNewExpiryDate", paramsArray);
			return NetRuntimeSystem.Convert.ToDateTime(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingFirstName()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingFirstName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingFirstName(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingFirstName", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingLastName()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingLastName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingLastName(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingLastName", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingPhone()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingPhone", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingPhone(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingPhone", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingAddress1()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingAddress1", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingAddress1(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingAddress1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingAddress2()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingAddress2", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingAddress2(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingAddress2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingCity()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingCity", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingCity(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingCity", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingState()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingState", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingState(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingState", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingCountryCode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingCountryCode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingCountryCode(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingCountryCode", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBillingZip()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBillingZip", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrNewVal">string bstrNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetBillingZip(string bstrNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrNewVal);
			Invoker.Method(this, "SetBillingZip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bSave">Int32 bSave</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 SaveBillingInfo(Int32 bSave)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bSave);
			object returnItem = Invoker.MethodReturn(this, "SaveBillingInfo", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCountryCode">string bstrCountryCode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 IsCCRenewalCountry(string bstrCountryCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCountryCode);
			object returnItem = Invoker.MethodReturn(this, "IsCCRenewalCountry", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCountryCode">string bstrCountryCode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetVATLabel(string bstrCountryCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCountryCode);
			object returnItem = Invoker.MethodReturn(this, "GetVATLabel", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public DateTime GetCCRenewalExpiryDate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCCRenewalExpiryDate", paramsArray);
			return NetRuntimeSystem.Convert.ToDateTime(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrVATNumber">string bstrVATNumber</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetVATNumber(string bstrVATNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrVATNumber);
			Invoker.Method(this, "SetVATNumber", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCCCode">string bstrCCCode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCreditCardType(string bstrCCCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCCCode);
			Invoker.Method(this, "SetCreditCardType", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCCNumber">string bstrCCNumber</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCreditCardNumber(string bstrCCNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCCNumber);
			Invoker.Method(this, "SetCreditCardNumber", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwCCYear">Int32 dwCCYear</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCreditCardExpiryYear(Int32 dwCCYear)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCCYear);
			Invoker.Method(this, "SetCreditCardExpiryYear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwCCMonth">Int32 dwCCMonth</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCreditCardExpiryMonth(Int32 dwCCMonth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCCMonth);
			Invoker.Method(this, "SetCreditCardExpiryMonth", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetCreditCardCount()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardCount", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCreditCardCode(Int32 dwIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwIndex);
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardCode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCreditCardName(Int32 dwIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwIndex);
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetVATNumber()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetVATNumber", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCreditCardType()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardType", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCreditCardNumber()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardNumber", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetCreditCardExpiryYear()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardExpiryYear", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetCreditCardExpiryMonth()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCreditCardExpiryMonth", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetDisconnectOption()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetDisconnectOption", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bNewVal">Int32 bNewVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetDisconnectOption(Int32 bNewVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bNewVal);
			Invoker.Method(this, "SetDisconnectOption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bReviseCustInfo">Int32 bReviseCustInfo</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessHandshakeRequest(Int32 bReviseCustInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bReviseCustInfo);
			Invoker.Method(this, "AsyncProcessHandshakeRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessNewLicenseRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessNewLicenseRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessReissueLicenseRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessReissueLicenseRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessRetailRenewalLicenseRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessRetailRenewalLicenseRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessReviseCustInfoRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessReviseCustInfoRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessCCRenewalPriceRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessCCRenewalPriceRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessCCRenewalLicenseRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessCCRenewalLicenseRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetAsyncProcessReturnCode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAsyncProcessReturnCode", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 IsUpgradeAvailable()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "IsUpgradeAvailable", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bWantUpgrade">Int32 bWantUpgrade</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void WantUpgrade(Int32 bWantUpgrade)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bWantUpgrade);
			Invoker.Method(this, "WantUpgrade", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void AsyncProcessDroppedLicenseRequest()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AsyncProcessDroppedLicenseRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GenerateInstallationId()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GenerateInstallationId", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrVal">string bstrVal</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 DepositConfirmationId(string bstrVal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrVal);
			object returnItem = Invoker.MethodReturn(this, "DepositConfirmationId", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCIDIID">string bstrCIDIID</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 VerifyCheckDigits(string bstrCIDIID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCIDIID);
			object returnItem = Invoker.MethodReturn(this, "VerifyCheckDigits", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public DateTime GetCurrentExpiryDate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCurrentExpiryDate", paramsArray);
			return NetRuntimeSystem.Convert.ToDateTime(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bIsLicenseRequest">Int32 bIsLicenseRequest</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void CancelAsyncProcessRequest(Int32 bIsLicenseRequest)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bIsLicenseRequest);
			Invoker.Method(this, "CancelAsyncProcessRequest", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetCurrencyDescription(Int32 dwCurrencyIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCurrencyIndex);
			object returnItem = Invoker.MethodReturn(this, "GetCurrencyDescription", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetPriceItemCount()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetPriceItemCount", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetPriceItemLabel(Int32 dwIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwIndex);
			object returnItem = Invoker.MethodReturn(this, "GetPriceItemLabel", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwCurrencyIndex">Int32 dwCurrencyIndex</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetPriceItemValue(Int32 dwCurrencyIndex, Int32 dwIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCurrencyIndex, dwIndex);
			object returnItem = Invoker.MethodReturn(this, "GetPriceItemValue", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetInvoiceText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetInvoiceText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetBackendErrorMsg()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetBackendErrorMsg", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetCurrencyOption()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetCurrencyOption", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dwCurrencyOption">Int32 dwCurrencyOption</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetCurrencyOption(Int32 dwCurrencyOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCurrencyOption);
			Invoker.Method(this, "SetCurrencyOption", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string GetEndOfLifeHtmlText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetEndOfLifeHtmlText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 DisplaySSLCert()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DisplaySSLCert", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}