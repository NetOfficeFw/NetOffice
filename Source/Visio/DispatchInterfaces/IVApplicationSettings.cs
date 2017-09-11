using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVApplicationSettings 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVApplicationSettings : COMObject
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
                    _type = typeof(IVApplicationSettings);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVApplicationSettings(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVApplicationSettings(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVApplicationSettings(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVApplicationSettings(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVApplicationSettings(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVApplicationSettings(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVApplicationSettings() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVApplicationSettings(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisObjectTypes ObjectType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisObjectTypes>(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool DrawingAids
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DrawingAids");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawingAids", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthRulerX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthRulerX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthRulerX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthRulerY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthRulerY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthRulerY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthGridX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthGridX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthGridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthGridY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthGridY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthGridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthGuidesX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthGuidesX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthGuidesX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthGuidesY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthGuidesY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthGuidesY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthPointsX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthPointsX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthPointsX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthPointsY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthPointsY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthPointsY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthGeometryX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthGeometryX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthGeometryX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthGeometryY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthGeometryY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthGeometryY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthExtensionsX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthExtensionsX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthExtensionsX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 SnapStrengthExtensionsY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SnapStrengthExtensionsY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapStrengthExtensionsY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowFileSaveWarnings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFileSaveWarnings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFileSaveWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowFileOpenWarnings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFileOpenWarnings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFileOpenWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisDefaultSaveFormats DefaultSaveFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisDefaultSaveFormats>(this, "DefaultSaveFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultSaveFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 DrawingPageColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DrawingPageColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawingPageColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 DrawingBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DrawingBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawingBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 DrawingBackgroundColorGradient
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DrawingBackgroundColorGradient");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawingBackgroundColorGradient", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 StencilBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StencilBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StencilBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 StencilBackgroundColorGradient
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StencilBackgroundColorGradient");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StencilBackgroundColorGradient", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 StencilTextColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StencilTextColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StencilTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 PrintPreviewBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PrintPreviewBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintPreviewBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 FullScreenBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FullScreenBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FullScreenBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowStartupDialog
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowStartupDialog");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowStartupDialog", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowSmartTags
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSmartTags");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSmartTags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisTextDisplayQualityTypes TextDisplayQuality
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisTextDisplayQualityTypes>(this, "TextDisplayQuality");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextDisplayQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool HigherQualityShapeDisplay
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HigherQualityShapeDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HigherQualityShapeDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool SmoothDrawing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SmoothDrawing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmoothDrawing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 StencilCharactersPerLine
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StencilCharactersPerLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StencilCharactersPerLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 StencilLinesPerMaster
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StencilLinesPerMaster");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StencilLinesPerMaster", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string UserName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string UserInitials
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserInitials");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserInitials", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ZoomOnRoll
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ZoomOnRoll");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ZoomOnRoll", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 UndoLevels
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UndoLevels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UndoLevels", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 RecentFilesListSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecentFilesListSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecentFilesListSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool CenterSelectionOnZoom
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CenterSelectionOnZoom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CenterSelectionOnZoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ConnectorSplittingEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConnectorSplittingEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectorSplittingEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisRegionalUIOptions AsianTextUI
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "AsianTextUI");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AsianTextUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisRegionalUIOptions ComplexTextUI
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "ComplexTextUI");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ComplexTextUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisRegionalUIOptions KanaFindAndReplace
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "KanaFindAndReplace");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "KanaFindAndReplace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 FreeformDrawingPrecision
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FreeformDrawingPrecision");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FreeformDrawingPrecision", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 FreeformDrawingSmoothing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FreeformDrawingSmoothing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FreeformDrawingSmoothing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool DeveloperMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DeveloperMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeveloperMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool ShowChooseDrawingTypePane
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowChooseDrawingTypePane");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowChooseDrawingTypePane", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool ShowShapeSearchPane
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowShapeSearchPane");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowShapeSearchPane", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool ApplyThemesOnShapeAdd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyThemesOnShapeAdd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyThemesOnShapeAdd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisRegionalUIOptions SATextUI
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "SATextUI");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisRegionalUIOptions BIDITextUI
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "BIDITextUI");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisRegionalUIOptions KashidaTextUI
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "KashidaTextUI");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool ShowMoreShapeHandlesOnHover
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowMoreShapeHandlesOnHover");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowMoreShapeHandlesOnHover", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool EnableAutoConnect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableAutoConnect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableAutoConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool ApplyBackgroundToDocument
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyBackgroundToDocument");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyBackgroundToDocument", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool TransitionsEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TransitionsEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TransitionsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool EnableFormulaAutoComplete
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableFormulaAutoComplete");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableFormulaAutoComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool DeleteConnectorsEnabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DeleteConnectorsEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeleteConnectorsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 RecentTemplatesListSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecentTemplatesListSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecentTemplatesListSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportDataFormat RasterExportDataFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportDataFormat>(this, "RasterExportDataFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportDataFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportDataCompression RasterExportDataCompression
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportDataCompression>(this, "RasterExportDataCompression");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportDataCompression", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportColorReduction RasterExportColorReduction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportColorReduction>(this, "RasterExportColorReduction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportColorReduction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportColorFormat RasterExportColorFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportColorFormat>(this, "RasterExportColorFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportColorFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportOperation RasterExportOperation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportOperation>(this, "RasterExportOperation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportOperation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportRotation RasterExportRotation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportRotation>(this, "RasterExportRotation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportRotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.Enums.VisRasterExportFlip RasterExportFlip
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportFlip>(this, "RasterExportFlip");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RasterExportFlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 RasterExportBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RasterExportBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RasterExportBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 RasterExportTransparencyColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RasterExportTransparencyColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RasterExportTransparencyColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool RasterExportUseTransparencyColor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RasterExportUseTransparencyColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RasterExportUseTransparencyColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 RasterExportQuality
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RasterExportQuality");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RasterExportQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public NetOffice.VisioApi.Enums.VisSVGExportFormat SVGExportFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisSVGExportFormat>(this, "SVGExportFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SVGExportFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public bool EnableLowMemoryMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableLowMemoryMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableLowMemoryMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public bool EnterCommitsText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnterCommitsText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterCommitsText", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		/// <param name="resolutionUnits">optional NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits resolutionUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width, object height, object resolutionUnits)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportResolution", resolution, width, height, resolutionUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportResolution", resolution);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportResolution", resolution, width);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width, object height)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportResolution", resolution, width, height);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pResolution">NetOffice.VisioApi.Enums.VisRasterExportResolution pResolution</param>
		/// <param name="pWidth">Double pWidth</param>
		/// <param name="pHeight">Double pHeight</param>
		/// <param name="pResolutionUnits">NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits pResolutionUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void GetRasterExportResolution(out NetOffice.VisioApi.Enums.VisRasterExportResolution pResolution, out Double pWidth, out Double pHeight, out NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits pResolutionUnits)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pResolution = 0;
			pWidth = 0;
			pHeight = 0;
			pResolutionUnits = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pResolution, pWidth, pHeight, pResolutionUnits);
			Invoker.Method(this, "GetRasterExportResolution", paramsArray, modifiers);
			pResolution = (NetOffice.VisioApi.Enums.VisRasterExportResolution)paramsArray[0];
			pWidth = (Double)paramsArray[1];
			pHeight = (Double)paramsArray[2];
			pResolutionUnits = (NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		/// <param name="sizeUnits">optional NetOffice.VisioApi.Enums.VisRasterExportSizeUnits sizeUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width, object height, object sizeUnits)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportSize", size, width, height, sizeUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportSize", size);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportSize", size, width);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width, object height)
		{
			 Factory.ExecuteMethod(this, "SetRasterExportSize", size, width, height);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pSize">NetOffice.VisioApi.Enums.VisRasterExportSize pSize</param>
		/// <param name="pWidth">Double pWidth</param>
		/// <param name="pHeight">Double pHeight</param>
		/// <param name="pSizeUnits">NetOffice.VisioApi.Enums.VisRasterExportSizeUnits pSizeUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void GetRasterExportSize(out NetOffice.VisioApi.Enums.VisRasterExportSize pSize, out Double pWidth, out Double pHeight, out NetOffice.VisioApi.Enums.VisRasterExportSizeUnits pSizeUnits)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pSize = 0;
			pWidth = 0;
			pHeight = 0;
			pSizeUnits = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pSize, pWidth, pHeight, pSizeUnits);
			Invoker.Method(this, "GetRasterExportSize", paramsArray, modifiers);
			pSize = (NetOffice.VisioApi.Enums.VisRasterExportSize)paramsArray[0];
			pWidth = (Double)paramsArray[1];
			pHeight = (Double)paramsArray[2];
			pSizeUnits = (NetOffice.VisioApi.Enums.VisRasterExportSizeUnits)paramsArray[3];
		}

		#endregion

		#pragma warning restore
	}
}
