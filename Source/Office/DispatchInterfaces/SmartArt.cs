using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface SmartArt 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860804.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SmartArt : _IMsoDispObj
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
                    _type = typeof(SmartArt);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SmartArt(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SmartArt(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864691.aspx </remarks>
		[SupportByVersion("Office", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862828.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNodes AllNodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtNodes>(this, "AllNodes", NetOffice.OfficeApi.SmartArtNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860244.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNodes Nodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtNodes>(this, "Nodes", NetOffice.OfficeApi.SmartArtNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861866.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtLayout Layout
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtLayout>(this, "Layout", NetOffice.OfficeApi.SmartArtLayout.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Layout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862785.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtQuickStyle QuickStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtQuickStyle>(this, "QuickStyle", NetOffice.OfficeApi.SmartArtQuickStyle.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "QuickStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862120.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtColor Color
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtColor>(this, "Color", NetOffice.OfficeApi.SmartArtColor.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Color", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865245.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864968.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		#endregion

		#pragma warning restore
	}
}
