using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface Window 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838990.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Window : COMObject
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
                    _type = typeof(Window);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Window(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197003.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839086.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838879.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822152.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Pane ActivePane
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActivePane", paramsArray);
				NetOffice.WordApi.Pane newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Pane.LateBindingApiWrapperType) as NetOffice.WordApi.Pane;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835485.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838919.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Panes Panes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Panes", paramsArray);
				NetOffice.WordApi.Panes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Panes.LateBindingApiWrapperType) as NetOffice.WordApi.Panes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845511.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Selection Selection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Selection", paramsArray);
				NetOffice.WordApi.Selection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Selection.LateBindingApiWrapperType) as NetOffice.WordApi.Selection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834260.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Left", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Left", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193092.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Top", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Top", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845639.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Width", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835119.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Height", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834813.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Split
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Split", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Split", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839134.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 SplitVertical
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SplitVertical", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SplitVertical", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822965.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Caption", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Caption", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845378.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdWindowState WindowState
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WindowState", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdWindowState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WindowState", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195421.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayRulers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayRulers", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayRulers", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835761.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayVerticalRuler
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayVerticalRuler", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayVerticalRuler", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838505.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.View View
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "View", paramsArray);
				NetOffice.WordApi.View newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.View.LateBindingApiWrapperType) as NetOffice.WordApi.View;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197875.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdWindowType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdWindowType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192589.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window Next
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Next", paramsArray);
				NetOffice.WordApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Window.LateBindingApiWrapperType) as NetOffice.WordApi.Window;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196868.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window Previous
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Previous", paramsArray);
				NetOffice.WordApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Window.LateBindingApiWrapperType) as NetOffice.WordApi.Window;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835402.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 WindowNumber
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WindowNumber", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837323.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayVerticalScrollBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayVerticalScrollBar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayVerticalScrollBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837924.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayHorizontalScrollBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayHorizontalScrollBar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayHorizontalScrollBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192230.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single StyleAreaWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StyleAreaWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StyleAreaWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840897.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayScreenTips
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayScreenTips", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayScreenTips", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191789.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 HorizontalPercentScrolled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HorizontalPercentScrolled", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HorizontalPercentScrolled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844796.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 VerticalPercentScrolled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VerticalPercentScrolled", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "VerticalPercentScrolled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839774.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DocumentMap
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentMap", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DocumentMap", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822144.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Active
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Active", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 DocumentMapPercentWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentMapPercentWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DocumentMapPercentWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194852.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192140.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdIMEMode IMEMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IMEMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdIMEMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IMEMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195060.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 UsableWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UsableWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821418.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 UsableHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UsableHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838517.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool EnvelopeVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnvelopeVisible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnvelopeVisible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835783.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayRightRuler
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayRightRuler", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayRightRuler", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195605.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayLeftScrollBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayLeftScrollBar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayLeftScrollBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820939.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192623.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool Thumbnails
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Thumbnails", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Thumbnails", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197684.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdShowSourceDocuments ShowSourceDocuments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowSourceDocuments", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdShowSourceDocuments)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowSourceDocuments", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231484.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public Int32 Hwnd
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hwnd", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838523.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845707.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="routeDocument">optional object RouteDocument</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object routeDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, routeDocument);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845707.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845707.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		/// <param name="toRight">optional object ToRight</param>
		/// <param name="toLeft">optional object ToLeft</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down, object up, object toRight, object toLeft)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up, toRight, toLeft);
			Invoker.Method(this, "LargeScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LargeScroll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LargeScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down);
			Invoker.Method(this, "LargeScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down, object up)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up);
			Invoker.Method(this, "LargeScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193698.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		/// <param name="toRight">optional object ToRight</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LargeScroll(object down, object up, object toRight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up, toRight);
			Invoker.Method(this, "LargeScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		/// <param name="toRight">optional object ToRight</param>
		/// <param name="toLeft">optional object ToLeft</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down, object up, object toRight, object toLeft)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up, toRight, toLeft);
			Invoker.Method(this, "SmallScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SmallScroll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SmallScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down);
			Invoker.Method(this, "SmallScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down, object up)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up);
			Invoker.Method(this, "SmallScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193450.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		/// <param name="toRight">optional object ToRight</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SmallScroll(object down, object up, object toRight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up, toRight);
			Invoker.Method(this, "SmallScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840287.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window NewWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NewWindow", paramsArray);
			NetOffice.WordApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Window.LateBindingApiWrapperType) as NetOffice.WordApi.Window;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839516.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		/// <param name="up">optional object Up</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PageScroll(object down, object up)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down, up);
			Invoker.Method(this, "PageScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839516.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PageScroll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PageScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839516.aspx
		/// </summary>
		/// <param name="down">optional object Down</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PageScroll(object down)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(down);
			Invoker.Method(this, "PageScroll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838905.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SetFocus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetFocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192575.aspx
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object RangeFromPoint(Int32 x, Int32 y)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x, y);
			object returnItem = Invoker.MethodReturn(this, "RangeFromPoint", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836655.aspx
		/// </summary>
		/// <param name="obj">object obj</param>
		/// <param name="start">optional object Start</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ScrollIntoView(object obj, object start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(obj, start);
			Invoker.Method(this, "ScrollIntoView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836655.aspx
		/// </summary>
		/// <param name="obj">object obj</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ScrollIntoView(object obj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(obj);
			Invoker.Method(this, "ScrollIntoView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836626.aspx
		/// </summary>
		/// <param name="screenPixelsLeft">Int32 ScreenPixelsLeft</param>
		/// <param name="screenPixelsTop">Int32 ScreenPixelsTop</param>
		/// <param name="screenPixelsWidth">Int32 ScreenPixelsWidth</param>
		/// <param name="screenPixelsHeight">Int32 ScreenPixelsHeight</param>
		/// <param name="obj">object obj</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void GetPoint(out Int32 screenPixelsLeft, out Int32 screenPixelsTop, out Int32 screenPixelsWidth, out Int32 screenPixelsHeight, object obj)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true,false);
			screenPixelsLeft = 0;
			screenPixelsTop = 0;
			screenPixelsWidth = 0;
			screenPixelsHeight = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(screenPixelsLeft, screenPixelsTop, screenPixelsWidth, screenPixelsHeight, obj);
			Invoker.Method(this, "GetPoint", paramsArray, modifiers);
			screenPixelsLeft = (Int32)paramsArray[0];
			screenPixelsTop = (Int32)paramsArray[1];
			screenPixelsWidth = (Int32)paramsArray[2];
			screenPixelsHeight = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object PrintZoomPaperHeight</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197226.aspx
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object PrintZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="background">optional object Background</param>
		/// <param name="append">optional object Append</param>
		/// <param name="range">optional object Range</param>
		/// <param name="outputFileName">optional object OutputFileName</param>
		/// <param name="from">optional object From</param>
		/// <param name="to">optional object To</param>
		/// <param name="item">optional object Item</param>
		/// <param name="copies">optional object Copies</param>
		/// <param name="pages">optional object Pages</param>
		/// <param name="pageType">optional object PageType</param>
		/// <param name="printToFile">optional object PrintToFile</param>
		/// <param name="collate">optional object Collate</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void ToggleShowAllReviewers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ToggleShowAllReviewers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835142.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ToggleRibbon()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ToggleRibbon", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}