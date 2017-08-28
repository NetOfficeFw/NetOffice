using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface Signature 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861800.aspx </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Signature : _IMsoDispObj
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
                    _type = typeof(Signature);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Signature(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Signature(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Signature(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Signature(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Signature(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Signature(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Signature() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Signature(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string Signer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Signer");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string Issuer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Issuer");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public object ExpireDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ExpireDate");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool IsValid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsValid");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool AttachCertificate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AttachCertificate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AttachCertificate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860302.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool IsCertificateExpired
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCertificateExpired");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool IsCertificateRevoked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCertificateRevoked");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public object SignDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SignDate");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864952.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool IsSigned
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSigned");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864576.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.SignatureInfo Details
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureInfo>(this, "Details", NetOffice.OfficeApi.SignatureInfo.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862368.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool CanSetup
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanSetup");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863325.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSetup Setup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSetup>(this, "Setup", NetOffice.OfficeApi.SignatureSetup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862851.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool IsSignatureLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSignatureLine");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863032.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object SignatureLineShape
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "SignatureLineShape");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863133.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 SortHint
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SortHint");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864585.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx </remarks>
		/// <param name="varSigImg">optional object varSigImg</param>
		/// <param name="varDelSuggSigner">optional object varDelSuggSigner</param>
		/// <param name="varDelSuggSignerLine2">optional object varDelSuggSignerLine2</param>
		/// <param name="varDelSuggSignerEmail">optional object varDelSuggSignerEmail</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Sign(object varSigImg, object varDelSuggSigner, object varDelSuggSignerLine2, object varDelSuggSignerEmail)
		{
			 Factory.ExecuteMethod(this, "Sign", varSigImg, varDelSuggSigner, varDelSuggSignerLine2, varDelSuggSignerEmail);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Sign()
		{
			 Factory.ExecuteMethod(this, "Sign");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx </remarks>
		/// <param name="varSigImg">optional object varSigImg</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Sign(object varSigImg)
		{
			 Factory.ExecuteMethod(this, "Sign", varSigImg);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx </remarks>
		/// <param name="varSigImg">optional object varSigImg</param>
		/// <param name="varDelSuggSigner">optional object varDelSuggSigner</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Sign(object varSigImg, object varDelSuggSigner)
		{
			 Factory.ExecuteMethod(this, "Sign", varSigImg, varDelSuggSigner);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx </remarks>
		/// <param name="varSigImg">optional object varSigImg</param>
		/// <param name="varDelSuggSigner">optional object varDelSuggSigner</param>
		/// <param name="varDelSuggSignerLine2">optional object varDelSuggSignerLine2</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Sign(object varSigImg, object varDelSuggSigner, object varDelSuggSignerLine2)
		{
			 Factory.ExecuteMethod(this, "Sign", varSigImg, varDelSuggSigner, varDelSuggSignerLine2);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860855.aspx </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ShowDetails()
		{
			 Factory.ExecuteMethod(this, "ShowDetails");
		}

		#endregion

		#pragma warning restore
	}
}
