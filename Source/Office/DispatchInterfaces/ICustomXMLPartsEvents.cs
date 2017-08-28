using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface ICustomXMLPartsEvents 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ICustomXMLPartsEvents : COMObject
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
                    _type = typeof(ICustomXMLPartsEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ICustomXMLPartsEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ICustomXMLPartsEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartsEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartsEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartsEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartsEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartsEvents() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartsEvents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="newPart">NetOffice.OfficeApi.CustomXMLPart newPart</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void PartAfterAdd(NetOffice.OfficeApi.CustomXMLPart newPart)
		{
			 Factory.ExecuteMethod(this, "PartAfterAdd", newPart);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="oldPart">NetOffice.OfficeApi.CustomXMLPart oldPart</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void PartBeforeDelete(NetOffice.OfficeApi.CustomXMLPart oldPart)
		{
			 Factory.ExecuteMethod(this, "PartBeforeDelete", oldPart);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="part">NetOffice.OfficeApi.CustomXMLPart part</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void PartAfterLoad(NetOffice.OfficeApi.CustomXMLPart part)
		{
			 Factory.ExecuteMethod(this, "PartAfterLoad", part);
		}

		#endregion

		#pragma warning restore
	}
}
