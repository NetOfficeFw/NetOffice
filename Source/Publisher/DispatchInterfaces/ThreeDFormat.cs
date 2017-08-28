using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ThreeDFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ThreeDFormat : COMObject
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
                    _type = typeof(ThreeDFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ThreeDFormat(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ThreeDFormat(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ThreeDFormat(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ThreeDFormat(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ThreeDFormat(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ThreeDFormat(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ThreeDFormat() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ThreeDFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Depth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Depth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Depth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorFormat ExtrusionColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "ExtrusionColor", NetOffice.PublisherApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoExtrusionColorType ExtrusionColorType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoExtrusionColorType>(this, "ExtrusionColorType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ExtrusionColorType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Perspective
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Perspective");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Perspective", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection PresetExtrusionDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection>(this, "PresetExtrusionDirection");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetLightingDirection PresetLightingDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetLightingDirection>(this, "PresetLightingDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PresetLightingDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetLightingSoftness PresetLightingSoftness
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetLightingSoftness>(this, "PresetLightingSoftness");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PresetLightingSoftness", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetMaterial PresetMaterial
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetMaterial>(this, "PresetMaterial");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PresetMaterial", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat PresetThreeDFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat>(this, "PresetThreeDFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single RotationX
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RotationX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RotationX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single RotationY
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RotationY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RotationY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Visible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public NetOffice.OfficeApi.Enums.MsoBevelType BevelTopType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBevelType>(this, "BevelTopType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BevelTopType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public Single BevelTopInset
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BevelTopInset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BevelTopInset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public Single BevelTopDepth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BevelTopDepth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BevelTopDepth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public NetOffice.OfficeApi.Enums.MsoBevelType BevelBottomType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBevelType>(this, "BevelBottomType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BevelBottomType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public Single BevelBottomInset
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BevelBottomInset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BevelBottomInset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public Single BevelBottomDepth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BevelBottomDepth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BevelBottomDepth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public Single ContourWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ContourWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ContourWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public NetOffice.PublisherApi.ColorFormat ContourColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "ContourColor", NetOffice.PublisherApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public Single FieldOfView
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "FieldOfView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FieldOfView", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementRotationX(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationX", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementRotationY(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationY", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ResetRotation()
		{
			 Factory.ExecuteMethod(this, "ResetRotation");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetThreeDFormat">NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetThreeDFormat(NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat)
		{
			 Factory.ExecuteMethod(this, "SetThreeDFormat", presetThreeDFormat);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetExtrusionDirection">NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetExtrusionDirection(NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection)
		{
			 Factory.ExecuteMethod(this, "SetExtrusionDirection", presetExtrusionDirection);
		}

		#endregion

		#pragma warning restore
	}
}
