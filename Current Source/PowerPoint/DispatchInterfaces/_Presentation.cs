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
	/// DispatchInterface _Presentation 
	/// SupportByVersion PowerPoint, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(COMObject replacedObject) : base(replacedObject)
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
		public NetOffice.PowerPointApi._Master SlideMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi._Master TitleMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TitleMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi._Master NotesMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NotesMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi._Master HandoutMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HandoutMaster", paramsArray);
				NetOffice.PowerPointApi._Master newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Slides Slides
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Slides", paramsArray);
				NetOffice.PowerPointApi.Slides newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Slides.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Slides;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.PageSetup PageSetup
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSetup", paramsArray);
				NetOffice.PowerPointApi.PageSetup newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PageSetup.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PageSetup;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ColorSchemes ColorSchemes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorSchemes", paramsArray);
				NetOffice.PowerPointApi.ColorSchemes newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ColorSchemes.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ColorSchemes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ExtraColors ExtraColors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ExtraColors", paramsArray);
				NetOffice.PowerPointApi.ExtraColors newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ExtraColors.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ExtraColors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.SlideShowSettings SlideShowSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideShowSettings", paramsArray);
				NetOffice.PowerPointApi.SlideShowSettings newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SlideShowSettings.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SlideShowSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Fonts Fonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fonts", paramsArray);
				NetOffice.PowerPointApi.Fonts newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Fonts.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Fonts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.DocumentWindows Windows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Windows", paramsArray);
				NetOffice.PowerPointApi.DocumentWindows newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.DocumentWindows.LateBindingApiWrapperType) as NetOffice.PowerPointApi.DocumentWindows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Tags Tags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tags", paramsArray);
				NetOffice.PowerPointApi.Tags newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Tags.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Tags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Shape DefaultShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultShape", paramsArray);
				NetOffice.PowerPointApi.Shape newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public object BuiltInDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuiltInDocumentProperties", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public object CustomDocumentProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomDocumentProperties", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBProject", paramsArray);
				NetOffice.VBIDEApi.VBProject newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.PrintOptions PrintOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintOptions", paramsArray);
				NetOffice.PowerPointApi.PrintOptions newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PrintOptions.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PrintOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public object Container
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Container", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.SlideShowWindow SlideShowWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideShowWindow", paramsArray);
				NetOffice.PowerPointApi.SlideShowWindow newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SlideShowWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SlideShowWindow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandBars", paramsArray);
				NetOffice.OfficeApi.CommandBars newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBars;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.PublishObjects PublishObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PublishObjects", paramsArray);
				NetOffice.PowerPointApi.PublishObjects newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.PublishObjects.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PublishObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.WebOptions WebOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebOptions", paramsArray);
				NetOffice.PowerPointApi.WebOptions newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.WebOptions.LateBindingApiWrapperType) as NetOffice.PowerPointApi.WebOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLProject", paramsArray);
				NetOffice.OfficeApi.HTMLProject newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType) as NetOffice.OfficeApi.HTMLProject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.Designs Designs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Designs", paramsArray);
				NetOffice.PowerPointApi.Designs newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Designs.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Designs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Signatures", paramsArray);
				NetOffice.OfficeApi.SignatureSet newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType) as NetOffice.OfficeApi.SignatureSet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
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
		/// SupportByVersion PowerPoint 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Permission", paramsArray);
				NetOffice.OfficeApi.Permission newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Permission.LateBindingApiWrapperType) as NetOffice.OfficeApi.Permission;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SharedWorkspace", paramsArray);
				NetOffice.OfficeApi.SharedWorkspace newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType) as NetOffice.OfficeApi.SharedWorkspace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sync", paramsArray);
				NetOffice.OfficeApi.Sync newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Sync.LateBindingApiWrapperType) as NetOffice.OfficeApi.Sync;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentLibraryVersions", paramsArray);
				NetOffice.OfficeApi.DocumentLibraryVersions newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentLibraryVersions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContentTypeProperties", paramsArray);
				NetOffice.OfficeApi.MetaProperties newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType) as NetOffice.OfficeApi.MetaProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
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
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
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
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerPolicy", paramsArray);
				NetOffice.OfficeApi.ServerPolicy newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType) as NetOffice.OfficeApi.ServerPolicy;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentInspectors", paramsArray);
				NetOffice.OfficeApi.DocumentInspectors newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentInspectors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
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
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomXMLParts", paramsArray);
				NetOffice.OfficeApi.CustomXMLParts newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType) as NetOffice.OfficeApi.CustomXMLParts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
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
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.PowerPointApi.CustomerData CustomerData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomerData", paramsArray);
				NetOffice.PowerPointApi.CustomerData newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.CustomerData.LateBindingApiWrapperType) as NetOffice.PowerPointApi.CustomerData;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.PowerPointApi.Research Research
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Research", paramsArray);
				NetOffice.PowerPointApi.Research newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Research.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Research;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.SectionProperties SectionProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SectionProperties", paramsArray);
				NetOffice.PowerPointApi.SectionProperties newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SectionProperties.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SectionProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Coauthoring Coauthoring
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Coauthoring", paramsArray);
				NetOffice.PowerPointApi.Coauthoring newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Coauthoring.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Coauthoring;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Broadcast Broadcast
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Broadcast", paramsArray);
				NetOffice.PowerPointApi.Broadcast newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Broadcast.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Broadcast;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi._Master AddTitleMaster()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddTitleMaster", paramsArray);
			NetOffice.PowerPointApi._Master newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi._Master;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void ApplyTemplate(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "ApplyTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.DocumentWindow NewWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NewWindow", paramsArray);
			NetOffice.PowerPointApi.DocumentWindow newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.DocumentWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address, string subAddress, bool newWindow, bool addHistory, string extraInfo, NetOffice.OfficeApi.Enums.MsoExtraInfoMethod method, string headerInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address, string subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address, string subAddress, bool newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address, string subAddress, bool newWindow, bool addHistory)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address, string subAddress, bool newWindow, bool addHistory, string extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void FollowHyperlink(string address, string subAddress, bool newWindow, bool addHistory, string extraInfo, NetOffice.OfficeApi.Enums.MsoExtraInfoMethod method)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, newWindow, addHistory, extraInfo, method);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void AddToFavorites()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToFavorites", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Unused()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Unused", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		/// <param name="collate">optional NetOffice.OfficeApi.Enums.MsoTriState Collate = -99</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void PrintOut(Int32 from, Int32 to, string printToFile, Int32 copies, NetOffice.OfficeApi.Enums.MsoTriState collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void PrintOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void PrintOut(Int32 from)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void PrintOut(Int32 from, Int32 to)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void PrintOut(Int32 from, Int32 to, string printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void PrintOut(Int32 from, Int32 to, string printToFile, Int32 copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(from, to, printToFile, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Save()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Save", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SaveAs(string fileName, NetOffice.PowerPointApi.Enums.PpSaveAsFileType fileFormat, NetOffice.OfficeApi.Enums.MsoTriState embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, embedTrueTypeFonts);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SaveAs(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SaveAs(string fileName, NetOffice.PowerPointApi.Enums.PpSaveAsFileType fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SaveCopyAs(string fileName, NetOffice.PowerPointApi.Enums.PpSaveAsFileType fileFormat, NetOffice.OfficeApi.Enums.MsoTriState embedTrueTypeFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat, embedTrueTypeFonts);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SaveCopyAs(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SaveCopyAs(string fileName, NetOffice.PowerPointApi.Enums.PpSaveAsFileType fileFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, fileFormat);
			Invoker.Method(this, "SaveCopyAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="filterName">string FilterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		/// <param name="scaleHeight">optional Int32 ScaleHeight = 0</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Export(string path, string filterName, Int32 scaleWidth, Int32 scaleHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, filterName, scaleWidth, scaleHeight);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="filterName">string FilterName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Export(string path, string filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, filterName);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="filterName">string FilterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Export(string path, string filterName, Int32 scaleWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, filterName, scaleWidth);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="text">string Text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void SetUndoText(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "SetUndoText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void UpdateLinks()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateLinks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void WebPagePreview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "WebPagePreview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="cp">NetOffice.OfficeApi.Enums.MsoEncoding cp</param>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding cp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cp);
			Invoker.Method(this, "ReloadAs", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="isDesignTemplate">NetOffice.OfficeApi.Enums.MsoTriState IsDesignTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void MakeIntoTemplate(NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(isDesignTemplate);
			Invoker.Method(this, "MakeIntoTemplate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void sblt(string s)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(s);
			Invoker.Method(this, "sblt", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void Merge(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "Merge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void CheckIn(bool saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void CheckIn()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void CheckIn(bool saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void CheckIn(bool saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public bool CanCheckIn()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CanCheckIn", paramsArray);
			return (bool)returnItem;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		/// <param name="includeAttachment">optional object IncludeAttachment</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void SendForReview(string recipients, string subject, bool showMessage, object includeAttachment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage, includeAttachment);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void SendForReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void SendForReview(string recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void SendForReview(string recipients, string subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void SendForReview(string recipients, string subject, bool showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendForReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void ReplyWithChanges(bool showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(showMessage);
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void ReplyWithChanges()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReplyWithChanges", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void EndReview()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EndReview", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">optional string FileName = </param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void AddBaseline(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "AddBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void AddBaseline()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void RemoveBaseline()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// </summary>
		/// <param name="passwordEncryptionProvider">string PasswordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string PasswordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 PasswordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">bool PasswordEncryptionFileProperties</param>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, bool passwordEncryptionFileProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
			Invoker.Method(this, "SetPasswordEncryptionOptions", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = false</param>
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public void SendFaxOverInternet(string recipients, string subject, bool showMessage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject, showMessage);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public void SendFaxOverInternet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public void SendFaxOverInternet(string recipients)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 11,12,14)]
		public void SendFaxOverInternet(string recipients, string subject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recipients, subject);
			Invoker.Method(this, "SendFaxOverInternet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		/// <param name="afterSlide">bool AfterSlide</param>
		/// <param name="sectionTitle">string sectionTitle</param>
		/// <param name="newSectionIndex">Int32 newSectionIndex</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void NewSectionAfter(Int32 index, bool afterSlide, string sectionTitle, out Int32 newSectionIndex)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			newSectionIndex = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(index, afterSlide, sectionTitle, newSectionIndex);
			Invoker.Method(this, "NewSectionAfter", paramsArray, modifiers);
			newSectionIndex = (Int32)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void DeleteSection(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "DeleteSection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void DisableSections()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DisableSections", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public string sectionTitle(Int32 index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "sectionTitle", paramsArray);
			return (string)returnItem;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType Type</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void RemoveDocumentInformation(NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			Invoker.Method(this, "RemoveDocumentInformation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		/// <param name="versionType">optional object VersionType</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void CheckInWithVersion(bool saveChanges, object comments, object makePublic, object versionType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic, versionType);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void CheckInWithVersion()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void CheckInWithVersion(bool saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void CheckInWithVersion(bool saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional object MakePublic</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void CheckInWithVersion(bool saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckInWithVersion", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName, bool includeDocProperties, bool keepIRMSettings, bool docStructureTags, bool bitmapMissingFonts, bool useISO19005_1, object externalExporter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="path">string Path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType FixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName, bool includeDocProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName, bool includeDocProperties, bool keepIRMSettings)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName, bool includeDocProperties, bool keepIRMSettings, bool docStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName, bool includeDocProperties, bool keepIRMSettings, bool docStructureTags, bool bitmapMissingFonts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
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
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, NetOffice.PowerPointApi.Enums.PpFixedFormatIntent intent, NetOffice.OfficeApi.Enums.MsoTriState frameSlides, NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder handoutOrder, NetOffice.PowerPointApi.Enums.PpPrintOutputType outputType, NetOffice.OfficeApi.Enums.MsoTriState printHiddenSlides, NetOffice.PowerPointApi.PrintRange printRange, NetOffice.PowerPointApi.Enums.PpPrintRangeType rangeType, string slideShowName, bool includeDocProperties, bool keepIRMSettings, bool docStructureTags, bool bitmapMissingFonts, bool useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTasks", paramsArray);
			NetOffice.OfficeApi.WorkflowTasks newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTasks;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetWorkflowTemplates", paramsArray);
			NetOffice.OfficeApi.WorkflowTemplates newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType) as NetOffice.OfficeApi.WorkflowTemplates;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void LockServerFile()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "LockServerFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="themeName">string themeName</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void ApplyTheme(string themeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(themeName);
			Invoker.Method(this, "ApplyTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="slideLibraryUrl">string SlideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		/// <param name="useSlideOrder">optional bool UseSlideOrder = false</param>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void PublishSlides(string slideLibraryUrl, bool overwrite, bool useSlideOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideLibraryUrl, overwrite, useSlideOrder);
			Invoker.Method(this, "PublishSlides", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="slideLibraryUrl">string SlideLibraryUrl</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void PublishSlides(string slideLibraryUrl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideLibraryUrl);
			Invoker.Method(this, "PublishSlides", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		/// <param name="slideLibraryUrl">string SlideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void PublishSlides(string slideLibraryUrl, bool overwrite)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slideLibraryUrl, overwrite);
			Invoker.Method(this, "PublishSlides", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public void Convert()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Convert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="withPresentation">string withPresentation</param>
		/// <param name="baselinePresentation">string baselinePresentation</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void MergeWithBaseline(string withPresentation, string baselinePresentation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(withPresentation, baselinePresentation);
			Invoker.Method(this, "MergeWithBaseline", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void AcceptAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AcceptAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void RejectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RejectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void EnsureAllMediaUpgraded()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "EnsureAllMediaUpgraded", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void Convert2(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "Convert2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		/// <param name="quality">optional Int32 Quality = 85</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void CreateVideo(string fileName, bool useTimingsAndNarrations, Int32 defaultSlideDuration, Int32 vertResolution, Int32 framesPerSecond, Int32 quality)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond, quality);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void CreateVideo(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void CreateVideo(string fileName, bool useTimingsAndNarrations)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void CreateVideo(string fileName, bool useTimingsAndNarrations, Int32 defaultSlideDuration)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void CreateVideo(string fileName, bool useTimingsAndNarrations, Int32 defaultSlideDuration, Int32 vertResolution)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public void CreateVideo(string fileName, bool useTimingsAndNarrations, Int32 defaultSlideDuration, Int32 vertResolution, Int32 framesPerSecond)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond);
			Invoker.Method(this, "CreateVideo", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}