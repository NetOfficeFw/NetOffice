﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ThreeDFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Depth"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.ExtrusionColor"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ColorFormat ExtrusionColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ColorFormat>(this, "ExtrusionColor", NetOffice.WordApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.ExtrusionColorType"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Perspective"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetExtrusionDirection"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection PresetExtrusionDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection>(this, "PresetExtrusionDirection");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetLightingDirection"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetLightingSoftness"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetMaterial"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetThreeDFormat"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat PresetThreeDFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat>(this, "PresetThreeDFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.RotationX"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.RotationY"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Visible"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetLighting"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.Z"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.BevelTopType"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.BevelTopInset"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.BevelTopDepth"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.BevelBottomType"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.BevelBottomInset"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.BevelBottomDepth"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.PresetCamera"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetCamera PresetCamera
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetCamera>(this, "PresetCamera");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.RotationZ"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.ContourWidth"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.ContourColor"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ColorFormat ContourColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ColorFormat>(this, "ContourColor", NetOffice.WordApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.FieldOfView"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.ProjectText"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.LightAngle"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.IncrementRotationX"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void IncrementRotationX(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationX", increment);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.IncrementRotationY"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void IncrementRotationY(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationY", increment);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.ResetRotation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ResetRotation()
		{
			 Factory.ExecuteMethod(this, "ResetRotation");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.SetExtrusionDirection"/> </remarks>
		/// <param name="presetExtrusionDirection">NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetExtrusionDirection(NetOffice.OfficeApi.Enums.MsoPresetExtrusionDirection presetExtrusionDirection)
		{
			 Factory.ExecuteMethod(this, "SetExtrusionDirection", presetExtrusionDirection);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.SetThreeDFormat"/> </remarks>
		/// <param name="presetThreeDFormat">NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetThreeDFormat(NetOffice.OfficeApi.Enums.MsoPresetThreeDFormat presetThreeDFormat)
		{
			 Factory.ExecuteMethod(this, "SetThreeDFormat", presetThreeDFormat);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.SetPresetCamera"/> </remarks>
		/// <param name="presetCamera">NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void SetPresetCamera(NetOffice.OfficeApi.Enums.MsoPresetCamera presetCamera)
		{
			 Factory.ExecuteMethod(this, "SetPresetCamera", presetCamera);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.IncrementRotationZ"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void IncrementRotationZ(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationZ", increment);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.IncrementRotationHorizontal"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void IncrementRotationHorizontal(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationHorizontal", increment);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ThreeDFormat.IncrementRotationVertical"/> </remarks>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void IncrementRotationVertical(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotationVertical", increment);
		}

		#endregion

		#pragma warning restore
	}
}
