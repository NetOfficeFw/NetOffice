using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface SlideShowView 
	/// SupportByVersion PowerPoint, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SlideShowView : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SlideShowView(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Int32 Zoom
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Zoom", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Slide Slide
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Slide", paramsArray);
				NetOffice.PowerPointApi.Slide newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Slide;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Enums.PpSlideShowPointerType PointerType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PointerType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpSlideShowPointerType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PointerType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Enums.PpSlideShowState State
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "State", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpSlideShowState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "State", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.OfficeApi.Enums.MsoTriState AcceleratorsEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AcceleratorsEnabled", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AcceleratorsEnabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Single PresentationElapsedTime
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PresentationElapsedTime", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Single SlideElapsedTime
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideElapsedTime", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SlideElapsedTime", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Slide LastSlideViewed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LastSlideViewed", paramsArray);
				NetOffice.PowerPointApi.Slide newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Slide.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Slide;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Enums.PpSlideShowAdvanceMode AdvanceMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AdvanceMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpSlideShowAdvanceMode)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ColorFormat PointerColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PointerColor", paramsArray);
				NetOffice.PowerPointApi.ColorFormat newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ColorFormat.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ColorFormat;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsNamedShow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsNamedShow", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public string SlideShowName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideShowName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Int32 CurrentShowPosition
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentShowPosition", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.OfficeApi.Enums.MsoTriState MediaControlsVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsVisible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Single MediaControlsLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Single MediaControlsTop
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsTop", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Single MediaControlsWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public Single MediaControlsHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MediaControlsHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="beginX">Single BeginX</param>
		/// <param name="beginY">Single BeginY</param>
		/// <param name="endX">Single EndX</param>
		/// <param name="endY">Single EndY</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void DrawLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(beginX, beginY, endX, endY);
			Invoker.Method(this, "DrawLine", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void EraseDrawing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EraseDrawing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void First()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "First", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Last()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Last", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Next()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Next", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Previous()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Previous", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="resetSlide">optional NetOffice.OfficeApi.Enums.MsoTriState ResetSlide = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void GotoSlide(Int32 index, NetOffice.OfficeApi.Enums.MsoTriState resetSlide)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, resetSlide);
			Invoker.Method(this, "GotoSlide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="index">Int32 index</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void GotoSlide(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "GotoSlide", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="slideShowName">string SlideShowName</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void GotoNamedShow(string slideShowName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideShowName);
			Invoker.Method(this, "GotoNamedShow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void EndNamedShow()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EndNamedShow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void ResetSlideTime()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetSlideTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Exit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Exit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="pTracker">NetOffice.PowerPointApi.MouseTracker pTracker</param>
		/// <param name="presenter">NetOffice.OfficeApi.Enums.MsoTriState Presenter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void InstallTracker(NetOffice.PowerPointApi.MouseTracker pTracker, NetOffice.OfficeApi.Enums.MsoTriState presenter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pTracker, presenter);
			Invoker.Method(this, "InstallTracker", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void GotoClick(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "GotoClick", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public Int32 GetClickIndex()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetClickIndex", paramsArray);
			return (Int32)returnItem;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public Int32 GetClickCount()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetClickCount", paramsArray);
			return (Int32)returnItem;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public bool FirstAnimationIsAutomatic()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FirstAnimationIsAutomatic", paramsArray);
			return (bool)returnItem;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="shapeId">object ShapeId</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Player Player(object shapeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shapeId);
			object returnItem = Invoker.MethodReturn(this, "Player", paramsArray);
			NetOffice.PowerPointApi.Player newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Player.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Player;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}