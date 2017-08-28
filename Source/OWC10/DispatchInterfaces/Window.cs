using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface Window 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Window : COMObject
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
                    _type = typeof(Window);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Window(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Window(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Workbook Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Workbook>(this, "Parent", NetOffice.OWC10Api.Workbook.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range ActiveCell
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "ActiveCell");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Pane ActivePane
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Pane>(this, "ActivePane", NetOffice.OWC10Api.Pane.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet ActiveSheet
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "ActiveSheet", NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Headings ColumnHeadings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Headings>(this, "ColumnHeadings", NetOffice.OWC10Api.Headings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayColumnHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayColumnHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayColumnHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayCustomHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayCustomHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayCustomHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayHorizontalScrollBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayHorizontalScrollBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayHorizontalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayRowHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayRowHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayRowHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayVerticalScrollBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayVerticalScrollBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayVerticalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayWorkbookTabs
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayWorkbookTabs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayWorkbookTabs", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayZeros
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayZeros");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayZeros", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool EnableResize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool FreezePanes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FreezePanes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FreezePanes", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GridlineColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridlineColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridlineColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlColorIndex GridlineColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlColorIndex>(this, "GridlineColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GridlineColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Panes Panes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Panes>(this, "Panes", NetOffice.OWC10Api.Panes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range RangeSelection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "RangeSelection");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Headings RowHeadings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Headings>(this, "RowHeadings", NetOffice.OWC10Api.Headings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ScrollColumn
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ScrollColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ScrollRow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ScrollRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Sheets SelectedSheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Sheets>(this, "SelectedSheets", NetOffice.OWC10Api.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Selection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Selection");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double TabRatio
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "TabRatio");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlWindowType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlWindowType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double UsableHeight
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "UsableHeight");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double UsableWidth
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "UsableWidth");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ViewableRange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ViewableRange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewableRange", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range VisibleRange
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "VisibleRange");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 WindowNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WindowNumber");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		/// <param name="toLeft">optional object toLeft</param>
		[SupportByVersion("OWC10", 1)]
		public void LargeScroll(object down, object up, object toRight, object toLeft)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LargeScroll()
		{
			 Factory.ExecuteMethod(this, "LargeScroll");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LargeScroll(object down)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LargeScroll(object down, object up)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void LargeScroll(object down, object up, object toRight)
		{
			 Factory.ExecuteMethod(this, "LargeScroll", down, up, toRight);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="points">Int32 points</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 PointsToScreenPixelsX(Int32 points)
		{
			return Factory.ExecuteInt32MethodGet(this, "PointsToScreenPixelsX", points);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="points">Int32 points</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 PointsToScreenPixelsY(Int32 points)
		{
			return Factory.ExecuteInt32MethodGet(this, "PointsToScreenPixelsY", points);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range RangeFromPoint(Int32 x, Int32 y)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "RangeFromPoint", x, y);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ResetHeadings()
		{
			 Factory.ExecuteMethod(this, "ResetHeadings");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("OWC10", 1)]
		public void ScrollIntoView(Int32 left, Int32 top, Int32 width, Int32 height, object start)
		{
			 Factory.ExecuteMethod(this, "ScrollIntoView", new object[]{ left, top, width, height, start });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ScrollIntoView(Int32 left, Int32 top, Int32 width, Int32 height)
		{
			 Factory.ExecuteMethod(this, "ScrollIntoView", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		/// <param name="toLeft">optional object toLeft</param>
		[SupportByVersion("OWC10", 1)]
		public void SmallScroll(object down, object up, object toRight, object toLeft)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down, up, toRight, toLeft);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SmallScroll()
		{
			 Factory.ExecuteMethod(this, "SmallScroll");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SmallScroll(object down)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SmallScroll(object down, object up)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down, up);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="down">optional object down</param>
		/// <param name="up">optional object up</param>
		/// <param name="toRight">optional object toRight</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SmallScroll(object down, object up, object toRight)
		{
			 Factory.ExecuteMethod(this, "SmallScroll", down, up, toRight);
		}

		#endregion

		#pragma warning restore
	}
}
