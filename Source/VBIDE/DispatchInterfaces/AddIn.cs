using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface AddIn 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AddIn : COMObject
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
                    _type = typeof(AddIn);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AddIn(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AddIn(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AddIn(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AddIn(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AddIn(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AddIn(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AddIn() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AddIn(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public string Description
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Description");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Description", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Addins Collection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Addins>(this, "Collection", NetOffice.VBIDEApi.Addins.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public string ProgId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProgId");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public string Guid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Guid");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public bool Connect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Connect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Connect", value);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3), ProxyResult]
		public object Object
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Object");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Object", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
