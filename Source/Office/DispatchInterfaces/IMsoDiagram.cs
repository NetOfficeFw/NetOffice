using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface IMsoDiagram 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IMsoDiagram : _IMsoDispObj
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
                    _type = typeof(IMsoDiagram);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMsoDiagram(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMsoDiagram(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoDiagram(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoDiagram(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoDiagram(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoDiagram(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoDiagram() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMsoDiagram(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
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
		public NetOffice.OfficeApi.DiagramNodes Nodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DiagramNodes>(this, "Nodes", NetOffice.OfficeApi.DiagramNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoDiagramType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoDiagramType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AutoLayout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AutoLayout");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Reverse
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Reverse");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Reverse", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AutoFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AutoFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoFormat", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Convert(NetOffice.OfficeApi.Enums.MsoDiagramType type)
		{
			 Factory.ExecuteMethod(this, "Convert", type);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void FitText()
		{
			 Factory.ExecuteMethod(this, "FitText");
		}

		#endregion

		#pragma warning restore
	}
}
