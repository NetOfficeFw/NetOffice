using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOAPPSETTINGS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOAPPSETTINGS : COMObject, NetOffice.VisioApi.LPVISIOAPPSETTINGS
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOAPPSETTINGS);
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
                    _type = typeof(LPVISIOAPPSETTINGS);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOAPPSETTINGS() : base()
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
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisObjectTypes ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisObjectTypes>(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool DrawingAids
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DrawingAids");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawingAids", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthRulerX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthRulerX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthRulerX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthRulerY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthRulerY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthRulerY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthGridX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthGridX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthGridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthGridY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthGridY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthGridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthGuidesX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthGuidesX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthGuidesX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthGuidesY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthGuidesY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthGuidesY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthPointsX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthPointsX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthPointsX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthPointsY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthPointsY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthPointsY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthGeometryX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthGeometryX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthGeometryX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthGeometryY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthGeometryY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthGeometryY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthExtensionsX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthExtensionsX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthExtensionsX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 SnapStrengthExtensionsY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SnapStrengthExtensionsY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapStrengthExtensionsY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ShowFileSaveWarnings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowFileSaveWarnings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowFileSaveWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ShowFileOpenWarnings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowFileOpenWarnings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowFileOpenWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisDefaultSaveFormats DefaultSaveFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisDefaultSaveFormats>(this, "DefaultSaveFormat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultSaveFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 DrawingPageColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DrawingPageColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawingPageColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 DrawingBackgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DrawingBackgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawingBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 DrawingBackgroundColorGradient
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DrawingBackgroundColorGradient");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawingBackgroundColorGradient", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 StencilBackgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StencilBackgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StencilBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 StencilBackgroundColorGradient
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StencilBackgroundColorGradient");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StencilBackgroundColorGradient", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 StencilTextColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StencilTextColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StencilTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 PrintPreviewBackgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PrintPreviewBackgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPreviewBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 FullScreenBackgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FullScreenBackgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FullScreenBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ShowStartupDialog
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowStartupDialog");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowStartupDialog", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ShowSmartTags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowSmartTags");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowSmartTags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisTextDisplayQualityTypes TextDisplayQuality
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisTextDisplayQualityTypes>(this, "TextDisplayQuality");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextDisplayQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool HigherQualityShapeDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HigherQualityShapeDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HigherQualityShapeDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool SmoothDrawing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SmoothDrawing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SmoothDrawing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 StencilCharactersPerLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StencilCharactersPerLine");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StencilCharactersPerLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 StencilLinesPerMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StencilLinesPerMaster");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StencilLinesPerMaster", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string UserName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string UserInitials
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserInitials");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserInitials", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ZoomOnRoll
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ZoomOnRoll");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ZoomOnRoll", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 UndoLevels
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UndoLevels");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UndoLevels", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 RecentFilesListSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecentFilesListSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecentFilesListSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool CenterSelectionOnZoom
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CenterSelectionOnZoom");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CenterSelectionOnZoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ConnectorSplittingEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ConnectorSplittingEnabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectorSplittingEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRegionalUIOptions AsianTextUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "AsianTextUI");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AsianTextUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRegionalUIOptions ComplexTextUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "ComplexTextUI");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ComplexTextUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRegionalUIOptions KanaFindAndReplace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "KanaFindAndReplace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "KanaFindAndReplace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 FreeformDrawingPrecision
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FreeformDrawingPrecision");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FreeformDrawingPrecision", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 FreeformDrawingSmoothing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FreeformDrawingSmoothing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FreeformDrawingSmoothing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool DeveloperMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DeveloperMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeveloperMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool ShowChooseDrawingTypePane
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowChooseDrawingTypePane");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowChooseDrawingTypePane", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual bool ShowShapeSearchPane
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowShapeSearchPane");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowShapeSearchPane", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual bool ApplyThemesOnShapeAdd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyThemesOnShapeAdd");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyThemesOnShapeAdd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRegionalUIOptions SATextUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "SATextUI");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRegionalUIOptions BIDITextUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "BIDITextUI");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRegionalUIOptions KashidaTextUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRegionalUIOptions>(this, "KashidaTextUI");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual bool ShowMoreShapeHandlesOnHover
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowMoreShapeHandlesOnHover");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowMoreShapeHandlesOnHover", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual bool EnableAutoConnect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableAutoConnect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableAutoConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool ApplyBackgroundToDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyBackgroundToDocument");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyBackgroundToDocument", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool TransitionsEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TransitionsEnabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TransitionsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool EnableFormulaAutoComplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableFormulaAutoComplete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableFormulaAutoComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool DeleteConnectorsEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DeleteConnectorsEnabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeleteConnectorsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 RecentTemplatesListSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecentTemplatesListSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecentTemplatesListSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportDataFormat RasterExportDataFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportDataFormat>(this, "RasterExportDataFormat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportDataFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportDataCompression RasterExportDataCompression
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportDataCompression>(this, "RasterExportDataCompression");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportDataCompression", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportColorReduction RasterExportColorReduction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportColorReduction>(this, "RasterExportColorReduction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportColorReduction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportColorFormat RasterExportColorFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportColorFormat>(this, "RasterExportColorFormat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportColorFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportOperation RasterExportOperation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportOperation>(this, "RasterExportOperation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportOperation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportRotation RasterExportRotation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportRotation>(this, "RasterExportRotation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportRotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisRasterExportFlip RasterExportFlip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisRasterExportFlip>(this, "RasterExportFlip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RasterExportFlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 RasterExportBackgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RasterExportBackgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RasterExportBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 RasterExportTransparencyColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RasterExportTransparencyColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RasterExportTransparencyColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual bool RasterExportUseTransparencyColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RasterExportUseTransparencyColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RasterExportUseTransparencyColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32 RasterExportQuality
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RasterExportQuality");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RasterExportQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual NetOffice.VisioApi.Enums.VisSVGExportFormat SVGExportFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisSVGExportFormat>(this, "SVGExportFormat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SVGExportFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual bool EnableLowMemoryMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableLowMemoryMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableLowMemoryMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual bool EnterCommitsText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnterCommitsText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterCommitsText", value);
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
		public virtual void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width, object height, object resolutionUnits)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportResolution", resolution, width, height, resolutionUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportResolution", resolution);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportResolution", resolution, width);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportResolution", resolution, width, height);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pResolution">NetOffice.VisioApi.Enums.VisRasterExportResolution pResolution</param>
		/// <param name="pWidth">Double pWidth</param>
		/// <param name="pHeight">Double pHeight</param>
		/// <param name="pResolutionUnits">NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits pResolutionUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void GetRasterExportResolution(out NetOffice.VisioApi.Enums.VisRasterExportResolution pResolution, out Double pWidth, out Double pHeight, out NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits pResolutionUnits)
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
		public virtual void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width, object height, object sizeUnits)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportSize", size, width, height, sizeUnits);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportSize", size);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportSize", size, width);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetRasterExportSize", size, width, height);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pSize">NetOffice.VisioApi.Enums.VisRasterExportSize pSize</param>
		/// <param name="pWidth">Double pWidth</param>
		/// <param name="pHeight">Double pHeight</param>
		/// <param name="pSizeUnits">NetOffice.VisioApi.Enums.VisRasterExportSizeUnits pSizeUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void GetRasterExportSize(out NetOffice.VisioApi.Enums.VisRasterExportSize pSize, out Double pWidth, out Double pHeight, out NetOffice.VisioApi.Enums.VisRasterExportSizeUnits pSizeUnits)
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

