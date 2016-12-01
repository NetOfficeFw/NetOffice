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
	/// DispatchInterface SignatureProvider 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861225.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SignatureProvider : COMObject
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
                    _type = typeof(SignatureProvider);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SignatureProvider(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureProvider(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureProvider(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureProvider(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureProvider(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureProvider() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureProvider(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861157.aspx
		/// </summary>
		/// <param name="siglnimg">NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object XmlDsigStream</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public stdole.Picture GenerateSignatureLineImage(NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(siglnimg, psigsetup, psiginfo, xmlDsigStream);
			object returnItem = Invoker.MethodReturn(this, "GenerateSignatureLineImage", paramsArray);
			stdole.Picture newObject = Factory.CreateObjectFromComProxy(this, returnItem) as stdole.Picture;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861424.aspx
		/// </summary>
		/// <param name="parentWindow">object ParentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ShowSignatureSetup(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parentWindow, psigsetup);
			Invoker.Method(this, "ShowSignatureSetup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864670.aspx
		/// </summary>
		/// <param name="parentWindow">object ParentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ShowSigningCeremony(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parentWindow, psigsetup, psiginfo);
			Invoker.Method(this, "ShowSigningCeremony", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864683.aspx
		/// </summary>
		/// <param name="queryContinue">object QueryContinue</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object XmlDsigStream</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void SignXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(queryContinue, psigsetup, psiginfo, xmlDsigStream);
			Invoker.Method(this, "SignXmlDsig", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860266.aspx
		/// </summary>
		/// <param name="parentWindow">object ParentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void NotifySignatureAdded(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parentWindow, psigsetup, psiginfo);
			Invoker.Method(this, "NotifySignatureAdded", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863028.aspx
		/// </summary>
		/// <param name="queryContinue">object QueryContinue</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object XmlDsigStream</param>
		/// <param name="pcontverres">NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres</param>
		/// <param name="pcertverres">NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void VerifyXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(queryContinue, psigsetup, psiginfo, xmlDsigStream, pcontverres, pcertverres);
			Invoker.Method(this, "VerifyXmlDsig", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865248.aspx
		/// </summary>
		/// <param name="parentWindow">object ParentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object XmlDsigStream</param>
		/// <param name="pcontverres">NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres</param>
		/// <param name="pcertverres">NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ShowSignatureDetails(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(parentWindow, psigsetup, psiginfo, xmlDsigStream, pcontverres, pcertverres);
			Invoker.Method(this, "ShowSignatureDetails", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863646.aspx
		/// </summary>
		/// <param name="sigprovdet">NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object GetProviderDetail(NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sigprovdet);
			object returnItem = Invoker.MethodReturn(this, "GetProviderDetail", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862104.aspx
		/// </summary>
		/// <param name="queryContinue">object QueryContinue</param>
		/// <param name="stream">object Stream</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public byte[] HashStream(object queryContinue, object stream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(queryContinue, stream);
			object returnItem = (object)Invoker.MethodReturn(this, "HashStream", paramsArray);
			return (byte[])returnItem;
		}

		#endregion
		#pragma warning restore
	}
}