using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface SignatureInfo 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865566.aspx </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SignatureInfo : _IMsoDispObj
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
                    _type = typeof(SignatureInfo);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SignatureInfo(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SignatureInfo(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureInfo(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureInfo(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureInfo(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureInfo(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureInfo() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureInfo(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860243.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool ReadOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865010.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string SignatureProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SignatureProvider");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860281.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string SignatureText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SignatureText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SignatureText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861498.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16), NativeResult]
		public stdole.Picture SignatureImage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SignatureImage", paramsArray);
                return returnItem as stdole.Picture;
            }
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SignatureImage", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860921.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string SignatureComment
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SignatureComment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SignatureComment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860572.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.ContentVerificationResults ContentVerificationResults
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.ContentVerificationResults>(this, "ContentVerificationResults");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864945.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.CertificateVerificationResults CertificateVerificationResults
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.CertificateVerificationResults>(this, "CertificateVerificationResults");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862453.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool IsValid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsValid");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860786.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool IsCertificateExpired
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCertificateExpired");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865218.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool IsCertificateRevoked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCertificateRevoked");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864566.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool IsCertificateUntrusted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCertificateUntrusted");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862539.aspx </remarks>
		/// <param name="sigdet">NetOffice.OfficeApi.Enums.SignatureDetail sigdet</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object GetSignatureDetail(NetOffice.OfficeApi.Enums.SignatureDetail sigdet)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetSignatureDetail", sigdet);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865451.aspx </remarks>
		/// <param name="certdet">NetOffice.OfficeApi.Enums.CertificateDetail certdet</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object GetCertificateDetail(NetOffice.OfficeApi.Enums.CertificateDetail certdet)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetCertificateDetail", certdet);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863087.aspx </remarks>
		/// <param name="parentWindow">object parentWindow</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ShowSignatureCertificate(object parentWindow)
		{
			 Factory.ExecuteMethod(this, "ShowSignatureCertificate", parentWindow);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863741.aspx </remarks>
		/// <param name="parentWindow">object parentWindow</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void SelectSignatureCertificate(object parentWindow)
		{
			 Factory.ExecuteMethod(this, "SelectSignatureCertificate", parentWindow);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863290.aspx </remarks>
		/// <param name="bstrThumbprint">string bstrThumbprint</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void SelectCertificateDetailByThumbprint(string bstrThumbprint)
		{
			 Factory.ExecuteMethod(this, "SelectCertificateDetailByThumbprint", bstrThumbprint);
		}

		#endregion

		#pragma warning restore
	}
}
