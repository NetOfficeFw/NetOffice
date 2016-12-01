using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface _Presentation 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Presentation : COMObject
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
                    _type = typeof(_Presentation);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Presentation(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745080.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743905.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745484.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi._Master SlideMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746378.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi._Master TitleMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TitleMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745657.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasTitleMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasTitleMaster", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744870.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public string TemplateName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TemplateName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743938.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi._Master NotesMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NotesMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746405.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi._Master HandoutMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HandoutMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746142.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slides Slides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Slides", paramsArray);
				NetOffice.PowerPointApi.Slides newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Slides.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Slides;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745413.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PageSetup PageSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSetup", paramsArray);
				NetOffice.PowerPointApi.PageSetup newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PageSetup.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PageSetup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744763.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorSchemes ColorSchemes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorSchemes", paramsArray);
				NetOffice.PowerPointApi.ColorSchemes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ColorSchemes.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ColorSchemes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746216.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ExtraColors ExtraColors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExtraColors", paramsArray);
				NetOffice.PowerPointApi.ExtraColors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ExtraColors.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ExtraColors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745621.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowSettings SlideShowSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideShowSettings", paramsArray);
				NetOffice.PowerPointApi.SlideShowSettings newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SlideShowSettings.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SlideShowSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744620.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Fonts Fonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fonts", paramsArray);
				NetOffice.PowerPointApi.Fonts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Fonts.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Fonts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746292.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindows Windows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Windows", paramsArray);
				NetOffice.PowerPointApi.DocumentWindows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.DocumentWindows.LateBindingApiWrapperType) as NetOffice.PowerPointApi.DocumentWindows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744602.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Tags Tags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tags", paramsArray);
				NetOffice.PowerPointApi.Tags newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Tags.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Tags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744397.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape DefaultShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultShape", paramsArray);
				NetOffice.PowerPointApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746376.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public object BuiltInDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuiltInDocumentProperties", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744661.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public object CustomDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomDocumentProperties", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745299.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBProject", paramsArray);
				NetOffice.VBIDEApi.VBProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746329.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ReadOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnly", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746313.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745890.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745125.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Path", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744884.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Saved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Saved", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Saved", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744109.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpDirection LayoutDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LayoutDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpDirection)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LayoutDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744540.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PrintOptions PrintOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintOptions", paramsArray);
				NetOffice.PowerPointApi.PrintOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PrintOptions.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PrintOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746277.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public object Container
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Container", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745979.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplayComments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayComments", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayComments", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746803.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel FarEastLineBreakLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FarEastLineBreakLevel", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FarEastLineBreakLevel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746404.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public string NoLineBreakBefore
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NoLineBreakBefore", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NoLineBreakBefore", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746110.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public string NoLineBreakAfter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NoLineBreakAfter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NoLineBreakAfter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745765.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowWindow SlideShowWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideShowWindow", paramsArray);
				NetOffice.PowerPointApi.SlideShowWindow newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SlideShowWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SlideShowWindow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746489.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FarEastLineBreakLanguage", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FarEastLineBreakLanguage", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745465.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID DefaultLanguageID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultLanguageID", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultLanguageID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746786.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PublishObjects PublishObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PublishObjects", paramsArray);
				NetOffice.PowerPointApi.PublishObjects newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PublishObjects.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PublishObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.WebOptions WebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebOptions", paramsArray);
				NetOffice.PowerPointApi.WebOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.WebOptions.LateBindingApiWrapperType) as NetOffice.PowerPointApi.WebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLProject", paramsArray);
				NetOffice.OfficeApi.HTMLProject newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType) as NetOffice.OfficeApi.HTMLProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746516.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState EnvelopeVisible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnvelopeVisible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnvelopeVisible", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746671.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState VBASigned
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBASigned", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746323.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState SnapToGrid
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapToGrid", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapToGrid", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744975.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public Single GridDistance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridDistance", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridDistance", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744959.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Designs Designs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Designs", paramsArray);
				NetOffice.PowerPointApi.Designs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Designs.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Designs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745705.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Signatures", paramsArray);
				NetOffice.OfficeApi.SignatureSet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType) as NetOffice.OfficeApi.SignatureSet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746134.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState RemovePersonalInformation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemovePersonalInformation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemovePersonalInformation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpRevisionInfo HasRevisionInfo
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasRevisionInfo", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpRevisionInfo)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743904.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745251.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionAlgorithm", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744792.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionKeyLength", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743937.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PasswordEncryptionFileProperties", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745703.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public string Password
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Password", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Password", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744704.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public string WritePassword
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WritePassword", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "WritePassword", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744658.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Permission", paramsArray);
				NetOffice.OfficeApi.Permission newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Permission.LateBindingApiWrapperType) as NetOffice.OfficeApi.Permission;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745343.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SharedWorkspace", paramsArray);
				NetOffice.OfficeApi.SharedWorkspace newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType) as NetOffice.OfficeApi.SharedWorkspace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745948.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sync", paramsArray);
				NetOffice.OfficeApi.Sync newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Sync.LateBindingApiWrapperType) as NetOffice.OfficeApi.Sync;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744818.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentLibraryVersions", paramsArray);
				NetOffice.OfficeApi.DocumentLibraryVersions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentLibraryVersions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745118.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContentTypeProperties", paramsArray);
				NetOffice.OfficeApi.MetaProperties newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType) as NetOffice.OfficeApi.MetaProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public Int32 SectionCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SectionCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public bool HasSections
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasSections", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745108.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerPolicy", paramsArray);
				NetOffice.OfficeApi.ServerPolicy newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType) as NetOffice.OfficeApi.ServerPolicy;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744654.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentInspectors", paramsArray);
				NetOffice.OfficeApi.DocumentInspectors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentInspectors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746792.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasVBProject", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745253.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLParts", paramsArray);
				NetOffice.OfficeApi.CustomXMLParts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLParts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743879.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public bool Final
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Final", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Final", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745858.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.CustomerData CustomerData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomerData", paramsArray);
				NetOffice.PowerPointApi.CustomerData newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.CustomerData.LateBindingApiWrapperType) as NetOffice.PowerPointApi.CustomerData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746518.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Research Research
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Research", paramsArray);
				NetOffice.PowerPointApi.Research newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Research.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Research;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745747.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public string EncryptionProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EncryptionProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EncryptionProvider", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744806.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.SectionProperties SectionProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SectionProperties", paramsArray);
				NetOffice.PowerPointApi.SectionProperties newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SectionProperties.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SectionProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745326.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Coauthoring Coauthoring
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Coauthoring", paramsArray);
				NetOffice.PowerPointApi.Coauthoring newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Coauthoring.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Coauthoring;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746423.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public bool InMergeMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InMergeMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744901.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Broadcast Broadcast
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Broadcast", paramsArray);
				NetOffice.PowerPointApi.Broadcast newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Broadcast.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Broadcast;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745775.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public bool HasNotesMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasNotesMaster", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744680.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public bool HasHandoutMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HasHandoutMaster", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743993.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpMediaTaskStatus CreateVideoStatus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CreateVideoStatus", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PowerPointApi.Enums.PpMediaTaskStatus)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229098.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public bool ChartDataPointTrack
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChartDataPointTrack", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ChartDataPointTrack", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229702.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Guides Guides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Guides", paramsArray);
				NetOffice.PowerPointApi.Guides newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Guides.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Guides;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746001.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi._Master AddTitleMaster()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddTitleMaster", paramsArray);
			NetOffice.PowerPointApi._Master newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743876.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void ApplyTemplate(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "ApplyTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744342.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindow NewWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NewWindow", paramsArray);
			NetOffice.PowerPointApi.DocumentWindow newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.DocumentWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744969.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToFavorites", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void Unused()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Unused", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		/// <param name="collate">optional NetOffice.OfficeApi.Enums.MsoTriState Collate = -99</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745194.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName, object fileFormat, object embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, embedTrueTypeFonts);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName, object fileFormat, object embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, embedTrueTypeFonts);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName, object fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="filterName">string FilterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		/// <param name="scaleHeight">optional Int32 ScaleHeight = 0</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName, object scaleWidth, object scaleHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, filterName, scaleWidth, scaleHeight);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="filterName">string FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, filterName);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="filterName">string FilterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName, object scaleWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, filterName, scaleWidth);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743857.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void SetUndoText(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "SetUndoText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744167.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void UpdateLinks()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateLinks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WebPagePreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="cp">NetOffice.OfficeApi.Enums.MsoEncoding cp</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding cp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cp);
			Invoker.Method(this, "ReloadAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="isDesignTemplate">NetOffice.OfficeApi.Enums.MsoTriState IsDesignTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void MakeIntoTemplate(NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(isDesignTemplate);
			Invoker.Method(this, "MakeIntoTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(s);
			Invoker.Method(this, "sblt", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228409.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void Merge(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744274.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public bool CanCheckIn()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CanCheckIn", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		/// <param name="includeAttachment">optional object IncludeAttachment</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage, includeAttachment);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showMessage);
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746226.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void EndReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EndReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">optional string FileName = </param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void AddBaseline(object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "AddBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void AddBaseline()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void RemoveBaseline()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743880.aspx
		/// </summary>
		/// <param name="passwordEncryptionProvider">string PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string PasswordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 PasswordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">bool PasswordEncryptionFileProperties</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, bool passwordEncryptionFileProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = false</param>
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		/// <param name="afterSlide">bool AfterSlide</param>
		/// <param name="sectionTitle">string sectionTitle</param>
		/// <param name="newSectionIndex">Int32 newSectionIndex</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void NewSectionAfter(Int32 index, bool afterSlide, string sectionTitle, out Int32 newSectionIndex)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			newSectionIndex = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(index, afterSlide, sectionTitle, newSectionIndex);
			Invoker.Method(this, "NewSectionAfter", paramsArray, modifiers);
			newSectionIndex = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void DeleteSection(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "DeleteSection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void DisableSections()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DisableSections", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public string sectionTitle(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "sectionTitle", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744345.aspx
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType Type</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "RemoveDocumentInformation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		/// <param name="versionType">optional object VersionType</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic, versionType);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="externalExporter">optional object ExternalExporter</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746373.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTasks", paramsArray);
			NetOffice.OfficeApi.WorkflowTasks newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTasks;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746712.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTemplates", paramsArray);
			NetOffice.OfficeApi.WorkflowTemplates newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTemplates;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744831.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void LockServerFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LockServerFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746528.aspx
		/// </summary>
		/// <param name="themeName">string themeName</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void ApplyTheme(string themeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(themeName);
			Invoker.Method(this, "ApplyTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx
		/// </summary>
		/// <param name="slideLibraryUrl">string SlideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		/// <param name="useSlideOrder">optional bool UseSlideOrder = false</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl, object overwrite, object useSlideOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideLibraryUrl, overwrite, useSlideOrder);
			Invoker.Method(this, "PublishSlides", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx
		/// </summary>
		/// <param name="slideLibraryUrl">string SlideLibraryUrl</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideLibraryUrl);
			Invoker.Method(this, "PublishSlides", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx
		/// </summary>
		/// <param name="slideLibraryUrl">string SlideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl, object overwrite)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideLibraryUrl, overwrite);
			Invoker.Method(this, "PublishSlides", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void Convert()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Convert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744063.aspx
		/// </summary>
		/// <param name="withPresentation">string withPresentation</param>
		/// <param name="baselinePresentation">string baselinePresentation</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void MergeWithBaseline(string withPresentation, string baselinePresentation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(withPresentation, baselinePresentation);
			Invoker.Method(this, "MergeWithBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745418.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void AcceptAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AcceptAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745993.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void RejectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RejectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744528.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void EnsureAllMediaUpgraded()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EnsureAllMediaUpgraded", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743830.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void Convert2(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "Convert2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		/// <param name="quality">optional Int32 Quality = 85</param>
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond, object quality)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond, quality);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228125.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="variant">string Variant</param>
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ApplyTemplate2(string fileName, string variant)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, variant);
			Invoker.Method(this, "ApplyTemplate2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="includeMarkup">optional bool IncludeMarkup = false</param>
		/// <param name="externalExporter">optional object ExternalExporter</param>
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object includeMarkup, object externalExporter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, includeMarkup, externalExporter);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="includeMarkup">optional bool IncludeMarkup = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object includeMarkup)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, includeMarkup);
			Invoker.Method(this, "ExportAsFixedFormat2", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}