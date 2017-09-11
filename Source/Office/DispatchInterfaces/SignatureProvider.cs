using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface SignatureProvider 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861225.aspx </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SignatureProvider : COMObject
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
                    _type = typeof(SignatureProvider);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SignatureProvider(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861157.aspx </remarks>
		/// <param name="siglnimg">NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object xmlDsigStream</param>
		[SupportByVersion("Office", 12,14,15,16), NativeResult]
		public stdole.Picture GenerateSignatureLineImage(NetOffice.OfficeApi.Enums.SignatureLineImage siglnimg, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(siglnimg, psigsetup, psiginfo, xmlDsigStream);
			object returnItem = Invoker.MethodReturn(this, "GenerateSignatureLineImage", paramsArray);
            return returnItem as stdole.Picture;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861424.aspx </remarks>
		/// <param name="parentWindow">object parentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ShowSignatureSetup(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup)
		{
			 Factory.ExecuteMethod(this, "ShowSignatureSetup", parentWindow, psigsetup);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864670.aspx </remarks>
		/// <param name="parentWindow">object parentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ShowSigningCeremony(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo)
		{
			 Factory.ExecuteMethod(this, "ShowSigningCeremony", parentWindow, psigsetup, psiginfo);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864683.aspx </remarks>
		/// <param name="queryContinue">object queryContinue</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object xmlDsigStream</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void SignXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream)
		{
			 Factory.ExecuteMethod(this, "SignXmlDsig", queryContinue, psigsetup, psiginfo, xmlDsigStream);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860266.aspx </remarks>
		/// <param name="parentWindow">object parentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void NotifySignatureAdded(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo)
		{
			 Factory.ExecuteMethod(this, "NotifySignatureAdded", parentWindow, psigsetup, psiginfo);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863028.aspx </remarks>
		/// <param name="queryContinue">object queryContinue</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object xmlDsigStream</param>
		/// <param name="pcontverres">NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres</param>
		/// <param name="pcertverres">NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void VerifyXmlDsig(object queryContinue, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres)
		{
			 Factory.ExecuteMethod(this, "VerifyXmlDsig", new object[]{ queryContinue, psigsetup, psiginfo, xmlDsigStream, pcontverres, pcertverres });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865248.aspx </remarks>
		/// <param name="parentWindow">object parentWindow</param>
		/// <param name="psigsetup">NetOffice.OfficeApi.SignatureSetup psigsetup</param>
		/// <param name="psiginfo">NetOffice.OfficeApi.SignatureInfo psiginfo</param>
		/// <param name="xmlDsigStream">object xmlDsigStream</param>
		/// <param name="pcontverres">NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres</param>
		/// <param name="pcertverres">NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ShowSignatureDetails(object parentWindow, NetOffice.OfficeApi.SignatureSetup psigsetup, NetOffice.OfficeApi.SignatureInfo psiginfo, object xmlDsigStream, NetOffice.OfficeApi.Enums.ContentVerificationResults pcontverres, NetOffice.OfficeApi.Enums.CertificateVerificationResults pcertverres)
		{
			 Factory.ExecuteMethod(this, "ShowSignatureDetails", new object[]{ parentWindow, psigsetup, psiginfo, xmlDsigStream, pcontverres, pcertverres });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863646.aspx </remarks>
		/// <param name="sigprovdet">NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object GetProviderDetail(NetOffice.OfficeApi.Enums.SignatureProviderDetail sigprovdet)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetProviderDetail", sigprovdet);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862104.aspx </remarks>
		/// <param name="queryContinue">object queryContinue</param>
		/// <param name="stream">object stream</param>
		[SupportByVersion("Office", 12,14,15,16)]
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
