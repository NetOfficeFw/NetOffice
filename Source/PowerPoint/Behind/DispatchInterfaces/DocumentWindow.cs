using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface DocumentWindow
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DocumentWindow : COMObject, NetOffice.PowerPointApi.DocumentWindow
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
                    _contractType = typeof(NetOffice.PowerPointApi.DocumentWindow);
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
                    _type = typeof(DocumentWindow);
                return _type;
            }
        }
        
        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DocumentWindow() : base()
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
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
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");			}
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Selection>(this, "Selection", typeof(NetOffice.PowerPointApi.Selection));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.View>(this, "View", typeof(NetOffice.PowerPointApi.View));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Presentation>(this, "Presentation", typeof(NetOffice.PowerPointApi.Presentation));
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
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpViewType>(this, "ViewType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ViewType", value);
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
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "BlackAndWhite");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BlackAndWhite", value);
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
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Active");
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
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpWindowState>(this, "WindowState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "WindowState", value);
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
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
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
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Left");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
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
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Top");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
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
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
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
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HWND");
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Pane>(this, "ActivePane", typeof(NetOffice.PowerPointApi.Pane));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Panes>(this, "Panes", typeof(NetOffice.PowerPointApi.Panes));
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SplitVertical");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SplitVertical", value);
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SplitHorizontal");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SplitHorizontal", value);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FitToPage");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down, up);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LargeScroll", down, up, toRight);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="down">optional Int32 Down = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down, up);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SmallScroll", down, up, toRight);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindow NewWindow()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DocumentWindow>(this, "NewWindow", typeof(NetOffice.PowerPointApi.DocumentWindow));
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public object RangeFromPoint(Int32 x, Int32 y)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RangeFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="points">Single points</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 PointsToScreenPixelsX(Single points)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PointsToScreenPixelsX", points);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="points">Single points</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 PointsToScreenPixelsY(Single points)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PointsToScreenPixelsY", points);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ScrollIntoView", new object[]{ left, top, width, height, start });
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ScrollIntoView", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool IsSectionExpanded(Int32 sectionIndex)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsSectionExpanded", sectionIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="sectionIndex">Int32 sectionIndex</param>
		/// <param name="expand">bool expand</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ExpandSection(Int32 sectionIndex, bool expand)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExpandSection", sectionIndex, expand);
		}

		#endregion

		#pragma warning restore
	}
}


