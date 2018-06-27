using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ThreeDFormat 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    public class ThreeDFormat : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ThreeDFormat
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.ThreeDFormat);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ThreeDFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single Depth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Depth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Depth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ColorFormat ExtrusionColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "ExtrusionColor", typeof(NetOffice.OfficeApi.ColorFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoExtrusionColorType ExtrusionColorType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoExtrusionColorType>(this, "ExtrusionColorType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ExtrusionColorType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState Perspective
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Perspective");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Perspective", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection PresetExtrusionDirection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection>(this, "PresetExtrusionDirection");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetLightingDirection PresetLightingDirection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetLightingDirection>(this, "PresetLightingDirection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PresetLightingDirection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetLightingSoftness PresetLightingSoftness
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetLightingSoftness>(this, "PresetLightingSoftness");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PresetLightingSoftness", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetMaterial PresetMaterial
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetMaterial>(this, "PresetMaterial");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PresetMaterial", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat PresetThreeDFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat>(this, "PresetThreeDFormat");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single RotationX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RotationX");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RotationX", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single RotationY
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RotationY");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RotationY", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState Visible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoLightRigType PresetLighting
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLightRigType>(this, "PresetLighting");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PresetLighting", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single Z
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Z");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Z", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoBevelType BevelTopType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBevelType>(this, "BevelTopType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BevelTopType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BevelTopInset
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BevelTopInset");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BevelTopInset", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BevelTopDepth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BevelTopDepth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BevelTopDepth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoBevelType BevelBottomType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBevelType>(this, "BevelBottomType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BevelBottomType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BevelBottomInset
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BevelBottomInset");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BevelBottomInset", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BevelBottomDepth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BevelBottomDepth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BevelBottomDepth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPresetCamera PresetCamera
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetCamera>(this, "PresetCamera");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single RotationZ
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RotationZ");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RotationZ", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single ContourWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ContourWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ContourWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ColorFormat ContourColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "ContourColor", typeof(NetOffice.OfficeApi.ColorFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single FieldOfView
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "FieldOfView");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FieldOfView", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState ProjectText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ProjectText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ProjectText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single LightAngle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LightAngle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LightAngle", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void IncrementRotationX(Single increment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "IncrementRotationX", increment);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void IncrementRotationY(Single increment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "IncrementRotationY", increment);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ResetRotation()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetRotation");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="presetThreeDFormat">NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SetThreeDFormat(NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetThreeDFormat", presetThreeDFormat);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="presetExtrusionDirection">NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SetExtrusionDirection(NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetExtrusionDirection", presetExtrusionDirection);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="presetCamera">NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetPresetCamera(NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPresetCamera", presetCamera);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void IncrementRotationZ(Single increment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "IncrementRotationZ", increment);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void IncrementRotationHorizontal(Single increment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "IncrementRotationHorizontal", increment);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void IncrementRotationVertical(Single increment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "IncrementRotationVertical", increment);
        }

        #endregion

        #pragma warning restore
    }
}
