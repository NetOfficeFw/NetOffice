using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface SlideShowView 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744675.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SlideShowView : COMObject
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
                    _type = typeof(SlideShowView);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SlideShowView(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SlideShowView(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746127.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744005.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745645.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Zoom
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Zoom");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744856.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slide Slide
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Slide>(this, "Slide", NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744941.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpSlideShowPointerType PointerType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpSlideShowPointerType>(this, "PointerType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PointerType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745272.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpSlideShowState State
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpSlideShowState>(this, "State");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "State", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743896.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AcceleratorsEnabled
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AcceleratorsEnabled");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AcceleratorsEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745215.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single PresentationElapsedTime
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PresentationElapsedTime");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746593.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single SlideElapsedTime
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SlideElapsedTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SlideElapsedTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744761.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slide LastSlideViewed
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Slide>(this, "LastSlideViewed", NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746290.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpSlideShowAdvanceMode AdvanceMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpSlideShowAdvanceMode>(this, "AdvanceMode");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744312.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorFormat PointerColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ColorFormat>(this, "PointerColor", NetOffice.PowerPointApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745856.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsNamedShow
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsNamedShow");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745076.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string SlideShowName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SlideShowName");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744575.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 CurrentShowPosition
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurrentShowPosition");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743996.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState MediaControlsVisible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "MediaControlsVisible");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744165.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Single MediaControlsLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MediaControlsLeft");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746549.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Single MediaControlsTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MediaControlsTop");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743865.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Single MediaControlsWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MediaControlsWidth");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744879.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Single MediaControlsHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MediaControlsHeight");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746368.aspx </remarks>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void DrawLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			 Factory.ExecuteMethod(this, "DrawLine", beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746340.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void EraseDrawing()
		{
			 Factory.ExecuteMethod(this, "EraseDrawing");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745022.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void First()
		{
			 Factory.ExecuteMethod(this, "First");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744036.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Last()
		{
			 Factory.ExecuteMethod(this, "Last");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746314.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Next()
		{
			 Factory.ExecuteMethod(this, "Next");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745842.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Previous()
		{
			 Factory.ExecuteMethod(this, "Previous");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746743.aspx </remarks>
		/// <param name="index">Int32 index</param>
		/// <param name="resetSlide">optional NetOffice.OfficeApi.Enums.MsoTriState ResetSlide = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void GotoSlide(Int32 index, object resetSlide)
		{
			 Factory.ExecuteMethod(this, "GotoSlide", index, resetSlide);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746743.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void GotoSlide(Int32 index)
		{
			 Factory.ExecuteMethod(this, "GotoSlide", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745384.aspx </remarks>
		/// <param name="slideShowName">string slideShowName</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void GotoNamedShow(string slideShowName)
		{
			 Factory.ExecuteMethod(this, "GotoNamedShow", slideShowName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744143.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void EndNamedShow()
		{
			 Factory.ExecuteMethod(this, "EndNamedShow");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745898.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ResetSlideTime()
		{
			 Factory.ExecuteMethod(this, "ResetSlideTime");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745742.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Exit()
		{
			 Factory.ExecuteMethod(this, "Exit");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTracker">NetOffice.PowerPointApi.MouseTracker pTracker</param>
		/// <param name="presenter">NetOffice.OfficeApi.Enums.MsoTriState presenter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void InstallTracker(NetOffice.PowerPointApi.MouseTracker pTracker, NetOffice.OfficeApi.Enums.MsoTriState presenter)
		{
			 Factory.ExecuteMethod(this, "InstallTracker", pTracker, presenter);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745994.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void GotoClick(Int32 index)
		{
			 Factory.ExecuteMethod(this, "GotoClick", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745129.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Int32 GetClickIndex()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetClickIndex");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744635.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Int32 GetClickCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetClickCount");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745143.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool FirstAnimationIsAutomatic()
		{
			return Factory.ExecuteBoolMethodGet(this, "FirstAnimationIsAutomatic");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746401.aspx </remarks>
		/// <param name="shapeId">object shapeId</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Player Player(object shapeId)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Player>(this, "Player", NetOffice.PowerPointApi.Player.LateBindingApiWrapperType, shapeId);
		}

		#endregion

		#pragma warning restore
	}
}
