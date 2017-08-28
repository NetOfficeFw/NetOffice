using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface TextFrame2 
	/// SupportByVersion PowerPoint, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745938.aspx </remarks>
	[SupportByVersion("PowerPoint", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TextFrame2 : COMObject
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
                    _type = typeof(TextFrame2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TextFrame2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TextFrame2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745455.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746553.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746188.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746697.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Single MarginBottom
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746466.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Single MarginLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746260.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Single MarginRight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744717.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Single MarginTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745236.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744106.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoHorizontalAnchor HorizontalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoHorizontalAnchor>(this, "HorizontalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HorizontalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746484.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoVerticalAnchor VerticalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoVerticalAnchor>(this, "VerticalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744719.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPathFormat PathFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPathFormat>(this, "PathFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PathFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744141.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoWarpFormat WarpFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoWarpFormat>(this, "WarpFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WarpFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745351.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetTextEffect WordArtFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetTextEffect>(this, "WordArtFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WordArtFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746152.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState WordWrap
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "WordWrap");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744773.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutoSize AutoSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutoSize>(this, "AutoSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744121.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.ThreeDFormat ThreeD
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ThreeDFormat>(this, "ThreeD", NetOffice.PowerPointApi.ThreeDFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744869.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasText
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasText");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744292.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 TextRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "TextRange", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746348.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.TextColumn2 Column
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextColumn2>(this, "Column", NetOffice.OfficeApi.TextColumn2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744383.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Ruler2 Ruler
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Ruler2>(this, "Ruler", NetOffice.OfficeApi.Ruler2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746350.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState NoTextRotation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "NoTextRotation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "NoTextRotation", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744756.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void DeleteText()
		{
			 Factory.ExecuteMethod(this, "DeleteText");
		}

		#endregion

		#pragma warning restore
	}
}
