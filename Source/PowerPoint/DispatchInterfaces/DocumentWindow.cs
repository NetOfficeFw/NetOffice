using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface DocumentWindow 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DocumentWindow : COMObject
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
                    _type = typeof(DocumentWindow);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DocumentWindow(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DocumentWindow(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentWindow(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentWindow(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentWindow(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentWindow(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentWindow() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentWindow(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
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
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Selection Selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Selection>(this, "Selection", NetOffice.PowerPointApi.Selection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.View View
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.View>(this, "View", NetOffice.PowerPointApi.View.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation Presentation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Presentation>(this, "Presentation", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpViewType ViewType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpViewType>(this, "ViewType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ViewType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState BlackAndWhite
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "BlackAndWhite");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BlackAndWhite", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Active
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Active");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpWindowState WindowState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpWindowState>(this, "WindowState");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single Left
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single Top
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single Width
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Single Height
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 HWND
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HWND");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Pane ActivePane
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Pane>(this, "ActivePane", NetOffice.PowerPointApi.Pane.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Panes Panes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Panes>(this, "Panes", NetOffice.PowerPointApi.Panes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 SplitVertical
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SplitVertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SplitVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 SplitHorizontal
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SplitHorizontal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SplitHorizontal", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FitToPage()
		{
			 Factory.ExecuteMethod(this, "FitToPage");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		/// <param name="toLeft">optional Int32 ToLeft = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down, object up, object toRight, object toLeft)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll()
		{
			 Factory.ExecuteMethod(this, "LargeScroll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down, object up)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down, object up, object toRight)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down, up, toRight);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		/// <param name="toLeft">optional Int32 ToLeft = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down, object up, object toRight, object toLeft)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll()
		{
			 Factory.ExecuteMethod(this, "SmallScroll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down, object up)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		/// <param name="up">optional Int32 Up = 0</param>
		/// <param name="toRight">optional Int32 ToRight = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down, object up, object toRight)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down, up, toRight);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindow NewWindow()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DocumentWindow>(this, "NewWindow", NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public object RangeFromPoint(Int32 x, Int32 y)
		{
			return Factory.ExecuteVariantMethodGet(this, "RangeFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="points">Single points</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 PointsToScreenPixelsX(Single points)
		{
			return Factory.ExecuteInt32MethodGet(this, "PointsToScreenPixelsX", points);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="points">Single points</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 PointsToScreenPixelsY(Single points)
		{
			return Factory.ExecuteInt32MethodGet(this, "PointsToScreenPixelsY", points);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="start">optional NetOffice.OfficeApi.Enums.MsoTriState Start = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ScrollIntoView(Single left, Single top, Single width, Single height, object start)
		{
			 Factory.ExecuteMethod(this, "ScrollIntoView", new object[]{ left, top, width, height, start });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ScrollIntoView(Single left, Single top, Single width, Single height)
		{
			 Factory.ExecuteMethod(this, "ScrollIntoView", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool IsSectionExpanded(Int32 sectionIndex)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsSectionExpanded", sectionIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="expand">bool expand</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ExpandSection(Int32 sectionIndex, bool expand)
		{
			 Factory.ExecuteMethod(this, "ExpandSection", sectionIndex, expand);
		}

		#endregion

		#pragma warning restore
	}
}
