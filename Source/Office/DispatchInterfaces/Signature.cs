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
	/// DispatchInterface Signature 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861800.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Signature : _IMsoDispObj
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
                    _type = typeof(Signature);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string Signer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Signer", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public string Issuer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Issuer", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public object ExpireDate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExpireDate", paramsArray);
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
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool IsValid
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsValid", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool AttachCertificate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AttachCertificate", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AttachCertificate", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860302.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool IsCertificateExpired
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsCertificateExpired", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool IsCertificateRevoked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsCertificateRevoked", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public object SignDate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SignDate", paramsArray);
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
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864952.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool IsSigned
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsSigned", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864576.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.SignatureInfo Details
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Details", paramsArray);
				NetOffice.OfficeApi.SignatureInfo newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SignatureInfo.LateBindingApiWrapperType) as NetOffice.OfficeApi.SignatureInfo;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862368.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool CanSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanSetup", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863325.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSetup Setup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Setup", paramsArray);
				NetOffice.OfficeApi.SignatureSetup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SignatureSetup.LateBindingApiWrapperType) as NetOffice.OfficeApi.SignatureSetup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862851.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool IsSignatureLine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsSignatureLine", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863032.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public object SignatureLineShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SignatureLineShape", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863133.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public Int32 SortHint
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SortHint", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864585.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx
		/// </summary>
		/// <param name="varSigImg">optional object varSigImg</param>
		/// <param name="varDelSuggSigner">optional object varDelSuggSigner</param>
		/// <param name="varDelSuggSignerLine2">optional object varDelSuggSignerLine2</param>
		/// <param name="varDelSuggSignerEmail">optional object varDelSuggSignerEmail</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Sign(object varSigImg, object varDelSuggSigner, object varDelSuggSignerLine2, object varDelSuggSignerEmail)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSigImg, varDelSuggSigner, varDelSuggSignerLine2, varDelSuggSignerEmail);
			Invoker.Method(this, "Sign", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Sign()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Sign", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx
		/// </summary>
		/// <param name="varSigImg">optional object varSigImg</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Sign(object varSigImg)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSigImg);
			Invoker.Method(this, "Sign", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx
		/// </summary>
		/// <param name="varSigImg">optional object varSigImg</param>
		/// <param name="varDelSuggSigner">optional object varDelSuggSigner</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Sign(object varSigImg, object varDelSuggSigner)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSigImg, varDelSuggSigner);
			Invoker.Method(this, "Sign", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861168.aspx
		/// </summary>
		/// <param name="varSigImg">optional object varSigImg</param>
		/// <param name="varDelSuggSigner">optional object varDelSuggSigner</param>
		/// <param name="varDelSuggSignerLine2">optional object varDelSuggSignerLine2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void Sign(object varSigImg, object varDelSuggSigner, object varDelSuggSignerLine2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varSigImg, varDelSuggSigner, varDelSuggSignerLine2);
			Invoker.Method(this, "Sign", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860855.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ShowDetails()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowDetails", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}