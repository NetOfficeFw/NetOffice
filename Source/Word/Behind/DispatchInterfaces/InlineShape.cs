using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface InlineShape 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836023.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class InlineShape : COMObject, NetOffice.WordApi.InlineShape
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
                    _contractType = typeof(NetOffice.WordApi.InlineShape);
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
                    _type = typeof(InlineShape);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public InlineShape() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844771.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838773.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193103.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834263.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Borders Borders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", typeof(NetOffice.WordApi.Borders));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Borders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840290.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Range
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839206.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.LinkFormat LinkFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LinkFormat>(this, "LinkFormat", typeof(NetOffice.WordApi.LinkFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838310.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field Field
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Field>(this, "Field", typeof(NetOffice.WordApi.Field));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193961.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.OLEFormat OLEFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OLEFormat>(this, "OLEFormat", typeof(NetOffice.WordApi.OLEFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845565.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdInlineShapeType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdInlineShapeType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194790.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Hyperlink Hyperlink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Hyperlink>(this, "Hyperlink", typeof(NetOffice.WordApi.Hyperlink));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192032.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single Height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837945.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838130.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single ScaleHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScaleHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197263.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single ScaleWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScaleWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195341.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState LockAspectRatio
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LockAspectRatio");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LockAspectRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838289.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.LineFormat Line
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LineFormat>(this, "Line", typeof(NetOffice.WordApi.LineFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838940.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.FillFormat Fill
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FillFormat>(this, "Fill", typeof(NetOffice.WordApi.FillFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193086.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.PictureFormat PictureFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PictureFormat>(this, "PictureFormat", typeof(NetOffice.WordApi.PictureFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "PictureFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194206.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.HorizontalLineFormat HorizontalLineFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HorizontalLineFormat>(this, "HorizontalLineFormat", typeof(NetOffice.WordApi.HorizontalLineFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834268.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Script Script
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Script>(this, "Script", typeof(NetOffice.OfficeApi.Script));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 OWSAnchor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "OWSAnchor");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193399.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TextEffectFormat TextEffect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TextEffectFormat>(this, "TextEffect", typeof(NetOffice.WordApi.TextEffectFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "TextEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836067.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string AlternativeText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AlternativeText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlternativeText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837925.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool IsPictureBullet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsPictureBullet");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837541.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.GroupShapes GroupItems
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.GroupShapes>(this, "GroupItems", typeof(NetOffice.WordApi.GroupShapes));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840752.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState HasChart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasChart");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool Dummy1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Dummy1");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835126.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.SoftEdgeFormat SoftEdge
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SoftEdgeFormat>(this, "SoftEdge", typeof(NetOffice.WordApi.SoftEdgeFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197808.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.GlowFormat Glow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.GlowFormat>(this, "Glow", typeof(NetOffice.WordApi.GlowFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197545.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ReflectionFormat Reflection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReflectionFormat>(this, "Reflection", typeof(NetOffice.WordApi.ReflectionFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836042.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.ShadowFormat Shadow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShadowFormat>(this, "Shadow", typeof(NetOffice.WordApi.ShadowFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193387.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.Chart Chart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Chart>(this, "Chart", typeof(NetOffice.WordApi.Chart));
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840914.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState HasSmartArt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasSmartArt");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840879.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.OfficeApi.SmartArt SmartArt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArt>(this, "SmartArt", typeof(NetOffice.OfficeApi.SmartArt));
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822328.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual string Title
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 AnchorID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AnchorID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 EditID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "EditID");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Activate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838092.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Reset()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839297.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839122.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Select()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193653.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shape ConvertToShape()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "ConvertToShape", typeof(NetOffice.WordApi.Shape));
		}

		#endregion

		#pragma warning restore
	}
}


