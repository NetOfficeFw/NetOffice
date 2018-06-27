using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface Font 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Font : COMObject, NetOffice.PublisherApi.Font
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
                    _contractType = typeof(NetOffice.PublisherApi.Font);
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
                    _type = typeof(Font);                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Font() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Bold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Bold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState BoldBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "BoldBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BoldBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Size");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object SizeBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SizeBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "SizeBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState AllCaps
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AllCaps");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AllCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Emboss
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Emboss");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Emboss", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Engrave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Engrave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Engrave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Italic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Italic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState ItalicBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ItalicBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ItalicBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Outline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Outline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Outline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState SmallCaps
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SmallCaps");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SmallCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState SuperScript
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SuperScript");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SuperScript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState SubScript
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SubScript");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SubScript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Shadow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Shadow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object AutomaticPairKerningThreshold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "AutomaticPairKerningThreshold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "AutomaticPairKerningThreshold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object Kerning
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Kerning");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Kerning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object Scaling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Scaling");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Scaling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object Tracking
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Tracking");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Tracking", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorFormat Color
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "Color", typeof(NetOffice.PublisherApi.ColorFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbTrackingPresetType TrackingPreset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbTrackingPresetType>(this, "TrackingPreset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TrackingPreset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbUnderlineType Underline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbUnderlineType>(this, "Underline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Underline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object Position
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Position");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual bool AttachedToText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AttachedToText");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState UseDiacriticColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "UseDiacriticColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "UseDiacriticColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorFormat DiacriticColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "DiacriticColor", typeof(NetOffice.PublisherApi.ColorFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState ExpandUsingKashida
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ExpandUsingKashida");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ExpandUsingKashida", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState Swash
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Swash");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Swash", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbNumberStylesType NumberStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbNumberStylesType>(this, "NumberStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "NumberStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbLigaturePresetType Ligature
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbLigaturePresetType>(this, "Ligature");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Ligature", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object StylisticAlternates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "StylisticAlternates");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "StylisticAlternates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState ContextualAlternates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ContextualAlternates");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ContextualAlternates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual object StylisticSets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "StylisticSets");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "StylisticSets", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState StrikeThrough
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "StrikeThrough");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "StrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public virtual NetOffice.PublisherApi.FillFormat Fill
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.FillFormat>(this, "Fill", typeof(NetOffice.PublisherApi.FillFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public virtual NetOffice.PublisherApi.LineFormat Line
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.LineFormat>(this, "Line", typeof(NetOffice.PublisherApi.LineFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public virtual NetOffice.PublisherApi.GlowFormat Glow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.GlowFormat>(this, "Glow", typeof(NetOffice.PublisherApi.GlowFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public virtual NetOffice.PublisherApi.ReflectionFormat Reflection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ReflectionFormat>(this, "Reflection", typeof(NetOffice.PublisherApi.ReflectionFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public virtual NetOffice.PublisherApi.ShadowFormat TextShadow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShadowFormat>(this, "TextShadow", typeof(NetOffice.PublisherApi.ShadowFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public virtual NetOffice.PublisherApi.ThreeDFormat ThreeD
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ThreeDFormat>(this, "ThreeD", typeof(NetOffice.PublisherApi.ThreeDFormat));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Grow()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Grow");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Shrink()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Shrink");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Font Duplicate()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Font>(this, "Duplicate", typeof(NetOffice.PublisherApi.Font));
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Reset()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="script">NetOffice.PublisherApi.Enums.PbFontScriptType script</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string GetScriptName(NetOffice.PublisherApi.Enums.PbFontScriptType script)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetScriptName", script);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="script">NetOffice.PublisherApi.Enums.PbFontScriptType script</param>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetScriptName(NetOffice.PublisherApi.Enums.PbFontScriptType script, string fontName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetScriptName", script, fontName);
		}

		#endregion

		#pragma warning restore
	}
}


