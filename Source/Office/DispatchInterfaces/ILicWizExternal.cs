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
	/// DispatchInterface ILicWizExternal 
	/// SupportByVersion Office, 10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ILicWizExternal : COMObject
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
                    _type = typeof(ILicWizExternal);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ILicWizExternal(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicWizExternal(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicWizExternal(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicWizExternal(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicWizExternal(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicWizExternal() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ILicWizExternal(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 Context
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Context", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public object Validator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Validator", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public object LicAgent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LicAgent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string CountryInfo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CountryInfo", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 WizardVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WizardVisible", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WizardVisible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string WizardTitle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WizardTitle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WizardTitle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 AnimationEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AnimationEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 CurrentHelpId
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentHelpId", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CurrentHelpId", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string OfficeOnTheWebUrl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OfficeOnTheWebUrl", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="punkHtmlDoc">object punkHtmlDoc</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void PrintHtmlDocument(object punkHtmlDoc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(punkHtmlDoc);
			Invoker.Method(this, "PrintHtmlDocument", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void InvokeDateTimeApplet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InvokeDateTimeApplet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="date">DateTime date</param>
		/// <param name="pFormat">optional string pFormat = </param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string FormatDate(DateTime date, object pFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(date, pFormat);
			object returnItem = Invoker.MethodReturn(this, "FormatDate", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="date">DateTime date</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string FormatDate(DateTime date)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(date);
			object returnItem = Invoker.MethodReturn(this, "FormatDate", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pvarId">optional object pvarId</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void ShowHelp(object pvarId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvarId);
			Invoker.Method(this, "ShowHelp", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void ShowHelp()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowHelp", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void Terminate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Terminate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bPC">Int32 BPC</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void DisableVORWReminder(Int32 bPC)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bPC);
			Invoker.Method(this, "DisableVORWReminder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrReceipt">string bstrReceipt</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string SaveReceipt(string bstrReceipt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrReceipt);
			object returnItem = Invoker.MethodReturn(this, "SaveReceipt", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrUrl">string bstrUrl</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void OpenInDefaultBrowser(string bstrUrl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrUrl);
			Invoker.Method(this, "OpenInDefaultBrowser", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrText">string bstrText</param>
		/// <param name="bstrButtons">string bstrButtons</param>
		/// <param name="bstrIcon">string bstrIcon</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 MsoAlert(string bstrText, string bstrButtons, string bstrIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrText, bstrButtons, bstrIcon);
			object returnItem = Invoker.MethodReturn(this, "MsoAlert", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrKey">string bstrKey</param>
		/// <param name="fMORW">Int32 fMORW</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 DepositPidKey(string bstrKey, Int32 fMORW)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrKey, fMORW);
			object returnItem = Invoker.MethodReturn(this, "DepositPidKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrMessage">string bstrMessage</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void WriteLog(string bstrMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrMessage);
			Invoker.Method(this, "WriteLog", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrProductCode">string bstrProductCode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void ResignDpc(string bstrProductCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrProductCode);
			Invoker.Method(this, "ResignDpc", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void ResetPID()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetPID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dx">Int32 dx</param>
		/// <param name="dy">Int32 dy</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SetDialogSize(Int32 dx, Int32 dy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dx, dy);
			Invoker.Method(this, "SetDialogSize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="lMode">Int32 lMode</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 VerifyClock(Int32 lMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lMode);
			object returnItem = Invoker.MethodReturn(this, "VerifyClock", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pdispSelect">object pdispSelect</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void SortSelectOptions(object pdispSelect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pdispSelect);
			Invoker.Method(this, "SortSelectOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void InternetDisconnect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "InternetDisconnect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public Int32 GetConnectedState()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetConnectedState", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}