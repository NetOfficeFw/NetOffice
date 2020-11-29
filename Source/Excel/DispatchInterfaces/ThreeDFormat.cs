using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ThreeDFormat 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.ThreeDFormat")]
    public class ThreeDFormat : NetOffice.OfficeApi._IMsoDispObj
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.Depth"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Single Depth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Depth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Depth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.ExtrusionColor"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.ColorFormat ExtrusionColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ColorFormat>(this, "ExtrusionColor", NetOffice.ExcelApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.ExtrusionColorType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.Perspective"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetExtrusionDirection"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection PresetExtrusionDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection>(this, "PresetExtrusionDirection");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetLightingDirection"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetLightingSoftness"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetMaterial"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetThreeDFormat"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat PresetThreeDFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat>(this, "PresetThreeDFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.RotationX"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.RotationY"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.Visible"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetLighting"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLightRigType PresetLighting
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLightRigType>(this, "PresetLighting");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PresetLighting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.Z"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Single Z
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Z");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Z", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.BevelTopType"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.BevelTopInset"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.BevelTopDepth"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.BevelBottomType"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.BevelBottomInset"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.BevelBottomDepth"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.PresetCamera"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetCamera PresetCamera
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetCamera>(this, "PresetCamera");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.RotationZ"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Single RotationZ
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RotationZ");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RotationZ", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.ContourWidth"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.ContourColor"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ColorFormat ContourColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ColorFormat>(this, "ContourColor", NetOffice.ExcelApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.FieldOfView"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.ProjectText"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ProjectText
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ProjectText");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ProjectText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.LightAngle"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Single LightAngle
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LightAngle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LightAngle", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.IncrementRotationX"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void IncrementRotationX(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationX", increment);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.IncrementRotationY"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void IncrementRotationY(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationY", increment);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.ResetRotation"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ResetRotation()
		{
			 Factory.ExecuteMethod(this, "ResetRotation");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.SetThreeDFormat"/> </remarks>
		/// <param name="presetThreeDFormat">NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SetThreeDFormat(NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat)
		{
			 Factory.ExecuteMethod(this, "SetThreeDFormat", presetThreeDFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.SetExtrusionDirection"/> </remarks>
		/// <param name="presetExtrusionDirection">NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SetExtrusionDirection(NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection)
		{
			 Factory.ExecuteMethod(this, "SetExtrusionDirection", presetExtrusionDirection);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.SetPresetCamera"/> </remarks>
		/// <param name="presetCamera">NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void SetPresetCamera(NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera)
		{
			 Factory.ExecuteMethod(this, "SetPresetCamera", presetCamera);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.IncrementRotationZ"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void IncrementRotationZ(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationZ", increment);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.IncrementRotationHorizontal"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void IncrementRotationHorizontal(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationHorizontal", increment);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.ThreeDFormat.IncrementRotationVertical"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void IncrementRotationVertical(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationVertical", increment);
		}

		#endregion

		#pragma warning restore
	}
}
