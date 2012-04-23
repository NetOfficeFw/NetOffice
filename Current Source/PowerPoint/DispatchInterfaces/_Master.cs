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
	/// DispatchInterface _Master 
	/// SupportByVersion PowerPoint, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Master : COMObject
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
                    _type = typeof(_Master);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Master(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Master(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Master(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Master() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Master(string progId) : base(progId)
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
				NetOffice.PowerPointApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
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
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Shapes Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				NetOffice.PowerPointApi.Shapes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Shapes.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Shapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.HeadersFooters HeadersFooters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeadersFooters", paramsArray);
				NetOffice.PowerPointApi.HeadersFooters newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.HeadersFooters.LateBindingApiWrapperType) as NetOffice.PowerPointApi.HeadersFooters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ColorScheme ColorScheme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColorScheme", paramsArray);
				NetOffice.PowerPointApi.ColorScheme newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ColorScheme.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ColorScheme;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ColorScheme", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.ShapeRange Background
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Background", paramsArray);
				NetOffice.PowerPointApi.ShapeRange newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ShapeRange;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get/Set
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
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Single Height
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Height", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public Single Width
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Width", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.TextStyles TextStyles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextStyles", paramsArray);
				NetOffice.PowerPointApi.TextStyles newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.TextStyles.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TextStyles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.PowerPointApi.Hyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.PowerPointApi.Hyperlinks newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Hyperlinks.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Hyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Scripts", paramsArray);
				NetOffice.OfficeApi.Scripts newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType) as NetOffice.OfficeApi.Scripts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.Design Design
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Design", paramsArray);
				NetOffice.PowerPointApi.Design newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Design.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Design;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.TimeLine TimeLine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TimeLine", paramsArray);
				NetOffice.PowerPointApi.TimeLine newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.TimeLine.LateBindingApiWrapperType) as NetOffice.PowerPointApi.TimeLine;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		public NetOffice.PowerPointApi.SlideShowTransition SlideShowTransition
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SlideShowTransition", paramsArray);
				NetOffice.PowerPointApi.SlideShowTransition newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.SlideShowTransition.LateBindingApiWrapperType) as NetOffice.PowerPointApi.SlideShowTransition;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.PowerPointApi.CustomLayouts CustomLayouts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomLayouts", paramsArray);
				NetOffice.PowerPointApi.CustomLayouts newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.CustomLayouts.LateBindingApiWrapperType) as NetOffice.PowerPointApi.CustomLayouts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.OfficeTheme Theme
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Theme", paramsArray);
				NetOffice.OfficeApi.OfficeTheme newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.OfficeTheme.LateBindingApiWrapperType) as NetOffice.OfficeApi.OfficeTheme;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14)]
		public NetOffice.OfficeApi.Enums.MsoBackgroundStyleIndex BackgroundStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackgroundStyle", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoBackgroundStyleIndex)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BackgroundStyle", paramsArray);
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
				NetOffice.PowerPointApi.CustomerData newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.CustomerData.LateBindingApiWrapperType) as NetOffice.PowerPointApi.CustomerData;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
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

		#endregion
		#pragma warning restore
	}
}