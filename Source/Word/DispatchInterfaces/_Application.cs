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
	/// DispatchInterface _Application 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Application : COMObject
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
                    _type = typeof(_Application);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Application(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823254.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197825.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191758.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845178.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821628.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Documents Documents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Documents", paramsArray);
				NetOffice.WordApi.Documents newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Documents.LateBindingApiWrapperType) as NetOffice.WordApi.Documents;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822351.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Windows Windows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Windows", paramsArray);
				NetOffice.WordApi.Windows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Windows.LateBindingApiWrapperType) as NetOffice.WordApi.Windows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837737.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document ActiveDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveDocument", paramsArray);
				NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845301.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window ActiveWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveWindow", paramsArray);
				NetOffice.WordApi.Window newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Window.LateBindingApiWrapperType) as NetOffice.WordApi.Window;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838682.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822917.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object WordBasic
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WordBasic", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195679.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.RecentFiles RecentFiles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecentFiles", paramsArray);
				NetOffice.WordApi.RecentFiles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.RecentFiles.LateBindingApiWrapperType) as NetOffice.WordApi.RecentFiles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845589.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Template NormalTemplate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NormalTemplate", paramsArray);
				NetOffice.WordApi.Template newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Template.LateBindingApiWrapperType) as NetOffice.WordApi.Template;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822391.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.System System
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "System", paramsArray);
				NetOffice.WordApi.System newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.System.LateBindingApiWrapperType) as NetOffice.WordApi.System;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845308.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCorrect AutoCorrect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoCorrect", paramsArray);
				NetOffice.WordApi.AutoCorrect newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.AutoCorrect.LateBindingApiWrapperType) as NetOffice.WordApi.AutoCorrect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197817.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FontNames FontNames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FontNames", paramsArray);
				NetOffice.WordApi.FontNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FontNames.LateBindingApiWrapperType) as NetOffice.WordApi.FontNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196340.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FontNames LandscapeFontNames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LandscapeFontNames", paramsArray);
				NetOffice.WordApi.FontNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FontNames.LateBindingApiWrapperType) as NetOffice.WordApi.FontNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192201.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FontNames PortraitFontNames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PortraitFontNames", paramsArray);
				NetOffice.WordApi.FontNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FontNames.LateBindingApiWrapperType) as NetOffice.WordApi.FontNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840701.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Languages Languages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Languages", paramsArray);
				NetOffice.WordApi.Languages newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Languages.LateBindingApiWrapperType) as NetOffice.WordApi.Languages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assistant", paramsArray);
				NetOffice.OfficeApi.Assistant newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType) as NetOffice.OfficeApi.Assistant;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821300.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Browser Browser
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Browser", paramsArray);
				NetOffice.WordApi.Browser newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Browser.LateBindingApiWrapperType) as NetOffice.WordApi.Browser;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823259.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FileConverters FileConverters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileConverters", paramsArray);
				NetOffice.WordApi.FileConverters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.FileConverters.LateBindingApiWrapperType) as NetOffice.WordApi.FileConverters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821659.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailingLabel MailingLabel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailingLabel", paramsArray);
				NetOffice.WordApi.MailingLabel newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.MailingLabel.LateBindingApiWrapperType) as NetOffice.WordApi.MailingLabel;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191745.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Dialogs Dialogs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dialogs", paramsArray);
				NetOffice.WordApi.Dialogs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Dialogs.LateBindingApiWrapperType) as NetOffice.WordApi.Dialogs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838479.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.CaptionLabels CaptionLabels
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CaptionLabels", paramsArray);
				NetOffice.WordApi.CaptionLabels newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.CaptionLabels.LateBindingApiWrapperType) as NetOffice.WordApi.CaptionLabels;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198063.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCaptions AutoCaptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoCaptions", paramsArray);
				NetOffice.WordApi.AutoCaptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.AutoCaptions.LateBindingApiWrapperType) as NetOffice.WordApi.AutoCaptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822986.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AddIns AddIns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AddIns", paramsArray);
				NetOffice.WordApi.AddIns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.AddIns.LateBindingApiWrapperType) as NetOffice.WordApi.AddIns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839544.aspx
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821519.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Version", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197438.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ScreenUpdating
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ScreenUpdating", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ScreenUpdating", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198164.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool PrintPreview
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintPreview", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintPreview", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839740.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tasks Tasks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tasks", paramsArray);
				NetOffice.WordApi.Tasks newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Tasks.LateBindingApiWrapperType) as NetOffice.WordApi.Tasks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayStatusBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayStatusBar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayStatusBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836086.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool SpecialMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SpecialMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839688.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834606.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192165.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool MathCoprocessorAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MathCoprocessorAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192426.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool MouseAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx
		/// </summary>
		/// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_International(NetOffice.WordApi.Enums.WdInternationalIndex index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "International", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx
		/// Alias for get_International
		/// </summary>
		/// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object International(NetOffice.WordApi.Enums.WdInternationalIndex index)
		{
			return get_International(index);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839495.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string Build
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Build", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820850.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CapsLock
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CapsLock", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845392.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool NumLock
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NumLock", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834599.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string UserName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844813.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string UserInitials
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserInitials", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserInitials", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193411.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string UserAddress
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserAddress", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserAddress", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835128.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object MacroContainer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MacroContainer", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838964.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayRecentFiles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayRecentFiles", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayRecentFiles", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845623.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="languageID">optional object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word, object languageID)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(word, languageID);
			object returnItem = Invoker.PropertyGet(this, "SynonymInfo", paramsArray);
			NetOffice.WordApi.SynonymInfo newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType) as NetOffice.WordApi.SynonymInfo;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx
		/// Alias for get_SynonymInfo
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="languageID">optional object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SynonymInfo SynonymInfo(string word, object languageID)
		{
			return get_SynonymInfo(word, languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(word);
			object returnItem = Invoker.PropertyGet(this, "SynonymInfo", paramsArray);
			NetOffice.WordApi.SynonymInfo newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType) as NetOffice.WordApi.SynonymInfo;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx
		/// Alias for get_SynonymInfo
		/// </summary>
		/// <param name="word">string Word</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SynonymInfo SynonymInfo(string word)
		{
			return get_SynonymInfo(word);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197234.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBE", paramsArray);
				NetOffice.VBIDEApi.VBE newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType) as NetOffice.VBIDEApi.VBE;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839412.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string DefaultSaveFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultSaveFormat", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultSaveFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821102.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListGalleries ListGalleries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ListGalleries", paramsArray);
				NetOffice.WordApi.ListGalleries newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ListGalleries.LateBindingApiWrapperType) as NetOffice.WordApi.ListGalleries;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821995.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ActivePrinter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActivePrinter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ActivePrinter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821925.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Templates Templates
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Templates", paramsArray);
				NetOffice.WordApi.Templates newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Templates.LateBindingApiWrapperType) as NetOffice.WordApi.Templates;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822548.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object CustomizationContext
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomizationContext", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CustomizationContext", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197596.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.KeyBindings KeyBindings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "KeyBindings", paramsArray);
				NetOffice.WordApi.KeyBindings newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.KeyBindings.LateBindingApiWrapperType) as NetOffice.WordApi.KeyBindings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx
		/// </summary>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory KeyCategory</param>
		/// <param name="command">string Command</param>
		/// <param name="commandParameter">optional object CommandParameter</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(keyCategory, command, commandParameter);
			object returnItem = Invoker.PropertyGet(this, "KeysBoundTo", paramsArray);
			NetOffice.WordApi.KeysBoundTo newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.KeysBoundTo.LateBindingApiWrapperType) as NetOffice.WordApi.KeysBoundTo;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx
		/// Alias for get_KeysBoundTo
		/// </summary>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory KeyCategory</param>
		/// <param name="command">string Command</param>
		/// <param name="commandParameter">optional object CommandParameter</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
		{
			return get_KeysBoundTo(keyCategory, command, commandParameter);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx
		/// </summary>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory KeyCategory</param>
		/// <param name="command">string Command</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(keyCategory, command);
			object returnItem = Invoker.PropertyGet(this, "KeysBoundTo", paramsArray);
			NetOffice.WordApi.KeysBoundTo newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.KeysBoundTo.LateBindingApiWrapperType) as NetOffice.WordApi.KeysBoundTo;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx
		/// Alias for get_KeysBoundTo
		/// </summary>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory KeyCategory</param>
		/// <param name="command">string Command</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
		{
			return get_KeysBoundTo(keyCategory, command);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx
		/// </summary>
		/// <param name="keyCode">Int32 KeyCode</param>
		/// <param name="keyCode2">optional object KeyCode2</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode, object keyCode2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(keyCode, keyCode2);
			object returnItem = Invoker.PropertyGet(this, "FindKey", paramsArray);
			NetOffice.WordApi.KeyBinding newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.KeyBinding.LateBindingApiWrapperType) as NetOffice.WordApi.KeyBinding;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx
		/// Alias for get_FindKey
		/// </summary>
		/// <param name="keyCode">Int32 KeyCode</param>
		/// <param name="keyCode2">optional object KeyCode2</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode, object keyCode2)
		{
			return get_FindKey(keyCode, keyCode2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx
		/// </summary>
		/// <param name="keyCode">Int32 KeyCode</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(keyCode);
			object returnItem = Invoker.PropertyGet(this, "FindKey", paramsArray);
			NetOffice.WordApi.KeyBinding newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.KeyBinding.LateBindingApiWrapperType) as NetOffice.WordApi.KeyBinding;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx
		/// Alias for get_FindKey
		/// </summary>
		/// <param name="keyCode">Int32 KeyCode</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode)
		{
			return get_FindKey(keyCode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196028.aspx
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
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192216.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192367.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayScrollBars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayScrollBars", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayScrollBars", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191937.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string StartupPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StartupPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StartupPath", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835146.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 BackgroundSavingStatus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackgroundSavingStatus", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820962.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 BackgroundPrintingStatus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackgroundPrintingStatus", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839318.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837463.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836284.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845159.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836388.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192152.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool DisplayAutoCompleteTips
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayAutoCompleteTips", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayAutoCompleteTips", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822542.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Options Options
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Options", paramsArray);
				NetOffice.WordApi.Options newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Options.LateBindingApiWrapperType) as NetOffice.WordApi.Options;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192373.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdAlertLevel DisplayAlerts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayAlerts", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdAlertLevel)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayAlerts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191957.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Dictionaries CustomDictionaries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomDictionaries", paramsArray);
				NetOffice.WordApi.Dictionaries newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Dictionaries.LateBindingApiWrapperType) as NetOffice.WordApi.Dictionaries;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192616.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string PathSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PathSeparator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845291.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string StatusBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StatusBar", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StatusBar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192800.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool MAPIAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MAPIAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845182.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839294.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdEnableCancelKey EnableCancelKey
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableCancelKey", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdEnableCancelKey)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableCancelKey", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197424.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool UserControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserControl", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileSearch", paramsArray);
				NetOffice.OfficeApi.FileSearch newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileSearch;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838972.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailSystem MailSystem
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailSystem", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdMailSystem)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839937.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string DefaultTableSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTableSeparator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultTableSeparator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839922.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool ShowVisualBasicEditor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowVisualBasicEditor", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowVisualBasicEditor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839549.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string BrowseExtraFileTypes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BrowseExtraFileTypes", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BrowseExtraFileTypes", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx
		/// </summary>
		/// <param name="_object">object Object</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsObjectValid(object _object)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(_object);
			object returnItem = Invoker.PropertyGet(this, "IsObjectValid", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx
		/// Alias for get_IsObjectValid
		/// </summary>
		/// <param name="_object">object Object</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool IsObjectValid(object _object)
		{
			return get_IsObjectValid(_object);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194713.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.HangulHanjaConversionDictionaries HangulHanjaDictionaries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HangulHanjaDictionaries", paramsArray);
				NetOffice.WordApi.HangulHanjaConversionDictionaries newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.HangulHanjaConversionDictionaries.LateBindingApiWrapperType) as NetOffice.WordApi.HangulHanjaConversionDictionaries;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821986.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMessage MailMessage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MailMessage", paramsArray);
				NetOffice.WordApi.MailMessage newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.MailMessage.LateBindingApiWrapperType) as NetOffice.WordApi.MailMessage;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840871.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool FocusInMailHeader
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FocusInMailHeader", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192588.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.EmailOptions EmailOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmailOptions", paramsArray);
				NetOffice.WordApi.EmailOptions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.EmailOptions.LateBindingApiWrapperType) as NetOffice.WordApi.EmailOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836711.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID Language
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Language", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoLanguageID)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192831.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "COMAddIns", paramsArray);
				NetOffice.OfficeApi.COMAddIns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType) as NetOffice.OfficeApi.COMAddIns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192428.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckLanguage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CheckLanguage", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CheckLanguage", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197161.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LanguageSettings", paramsArray);
				NetOffice.OfficeApi.LanguageSettings newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType) as NetOffice.OfficeApi.LanguageSettings;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy1
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dummy1", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AnswerWizard", paramsArray);
				NetOffice.OfficeApi.AnswerWizard newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType) as NetOffice.OfficeApi.AnswerWizard;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195192.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FeatureInstall", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFeatureInstall)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FeatureInstall", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192776.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutomationSecurity", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoAutomationSecurity)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutomationSecurity", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx
		/// </summary>
		/// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType FileDialogType</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fileDialogType);
			object returnItem = Invoker.PropertyGet(this, "FileDialog", paramsArray);
			NetOffice.OfficeApi.FileDialog newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType) as NetOffice.OfficeApi.FileDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx
		/// Alias for get_FileDialog
		/// </summary>
		/// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType FileDialogType</param>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
		{
			return get_FileDialog(fileDialogType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193382.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public string EmailTemplate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmailTemplate", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EmailTemplate", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool ShowWindowsInTaskbar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowWindowsInTaskbar", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowWindowsInTaskbar", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193065.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.NewFile NewDocument
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NewDocument", paramsArray);
				NetOffice.OfficeApi.NewFile newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType) as NetOffice.OfficeApi.NewFile;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840052.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool ShowStartupDialog
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowStartupDialog", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowStartupDialog", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192177.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCorrect AutoCorrectEmail
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoCorrectEmail", paramsArray);
				NetOffice.WordApi.AutoCorrect newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.AutoCorrect.LateBindingApiWrapperType) as NetOffice.WordApi.AutoCorrect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845341.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.TaskPanes TaskPanes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TaskPanes", paramsArray);
				NetOffice.WordApi.TaskPanes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.TaskPanes.LateBindingApiWrapperType) as NetOffice.WordApi.TaskPanes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835491.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool DefaultLegalBlackline
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultLegalBlackline", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultLegalBlackline", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.SmartTagRecognizers SmartTagRecognizers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTagRecognizers", paramsArray);
				NetOffice.WordApi.SmartTagRecognizers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SmartTagRecognizers.LateBindingApiWrapperType) as NetOffice.WordApi.SmartTagRecognizers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.SmartTagTypes SmartTagTypes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartTagTypes", paramsArray);
				NetOffice.WordApi.SmartTagTypes newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.SmartTagTypes.LateBindingApiWrapperType) as NetOffice.WordApi.SmartTagTypes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839771.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNamespaces XMLNamespaces
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLNamespaces", paramsArray);
				NetOffice.WordApi.XMLNamespaces newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLNamespaces.LateBindingApiWrapperType) as NetOffice.WordApi.XMLNamespaces;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196679.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public bool ArbitraryXMLSupportAvailable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ArbitraryXMLSupportAvailable", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BuildFull
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildFull", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BuildFeatureCrew
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildFeatureCrew", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192405.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Bibliography Bibliography
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Bibliography", paramsArray);
				NetOffice.WordApi.Bibliography newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Bibliography.LateBindingApiWrapperType) as NetOffice.WordApi.Bibliography;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191727.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool ShowStylePreviews
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowStylePreviews", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowStylePreviews", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845435.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool RestrictLinkedStyles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RestrictLinkedStyles", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RestrictLinkedStyles", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837322.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathAutoCorrect OMathAutoCorrect
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OMathAutoCorrect", paramsArray);
				NetOffice.WordApi.OMathAutoCorrect newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.OMathAutoCorrect.LateBindingApiWrapperType) as NetOffice.WordApi.OMathAutoCorrect;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836074.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool DisplayDocumentInformationPanel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayDocumentInformationPanel", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayDocumentInformationPanel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197133.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Assistance", paramsArray);
				NetOffice.OfficeApi.IAssistance newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType) as NetOffice.OfficeApi.IAssistance;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192620.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool OpenAttachmentsInFullScreen
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OpenAttachmentsInFullScreen", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OpenAttachmentsInFullScreen", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836063.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public Int32 ActiveEncryptionSession
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveEncryptionSession", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194203.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool DontResetInsertionPointProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DontResetInsertionPointProperties", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DontResetInsertionPointProperties", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839192.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartArtLayouts", paramsArray);
				NetOffice.OfficeApi.SmartArtLayouts newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtLayouts.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtLayouts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194982.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartArtQuickStyles", paramsArray);
				NetOffice.OfficeApi.SmartArtQuickStyles newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtQuickStyles.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtQuickStyles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839505.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtColors SmartArtColors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SmartArtColors", paramsArray);
				NetOffice.OfficeApi.SmartArtColors newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtColors.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtColors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838675.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.UndoRecord UndoRecord
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UndoRecord", paramsArray);
				NetOffice.WordApi.UndoRecord newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.UndoRecord.LateBindingApiWrapperType) as NetOffice.WordApi.UndoRecord;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191978.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.OfficeApi.PickerDialog PickerDialog
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PickerDialog", paramsArray);
				NetOffice.OfficeApi.PickerDialog newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.PickerDialog.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerDialog;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839925.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindows ProtectedViewWindows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectedViewWindows", paramsArray);
				NetOffice.WordApi.ProtectedViewWindows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ProtectedViewWindows.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192773.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow ActiveProtectedViewWindow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveProtectedViewWindow", paramsArray);
				NetOffice.WordApi.ProtectedViewWindow newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.WordApi.ProtectedViewWindow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845787.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public bool IsSandboxed
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsSandboxed", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193078.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileValidation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoFileValidationMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FileValidation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232091.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
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
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232207.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 15, 16)]
		public bool ShowAnimation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowAnimation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ShowAnimation", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="originalFormat">optional object OriginalFormat</param>
		/// <param name="routeDocument">optional object RouteDocument</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Quit(object saveChanges, object originalFormat, object routeDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, originalFormat, routeDocument);
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Quit(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx
		/// </summary>
		/// <param name="saveChanges">optional object SaveChanges</param>
		/// <param name="originalFormat">optional object OriginalFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Quit(object saveChanges, object originalFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, originalFormat);
			Invoker.Method(this, "Quit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193095.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ScreenRefresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ScreenRefresh", paramsArray);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint);
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
		/// <param name="fileName">optional object FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX);
			Invoker.Method(this, "PrintOutOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839803.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void LookupNameProperties(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			Invoker.Method(this, "LookupNameProperties", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192415.aspx
		/// </summary>
		/// <param name="unavailableFont">string UnavailableFont</param>
		/// <param name="substituteFont">string SubstituteFont</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SubstituteFont(string unavailableFont, string substituteFont)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unavailableFont, substituteFont);
			Invoker.Method(this, "SubstituteFont", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx
		/// </summary>
		/// <param name="times">optional object Times</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Repeat(object times)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(times);
			object returnItem = Invoker.MethodReturn(this, "Repeat", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Repeat()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Repeat", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845561.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		/// <param name="command">string Command</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DDEExecute(Int32 channel, string command)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel, command);
			Invoker.Method(this, "DDEExecute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837295.aspx
		/// </summary>
		/// <param name="app">string App</param>
		/// <param name="topic">string Topic</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 DDEInitiate(string app, string topic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(app, topic);
			object returnItem = Invoker.MethodReturn(this, "DDEInitiate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837201.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		/// <param name="item">string Item</param>
		/// <param name="data">string Data</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DDEPoke(Int32 channel, string item, string data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel, item, data);
			Invoker.Method(this, "DDEPoke", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837546.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		/// <param name="item">string Item</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string DDERequest(Int32 channel, string item)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel, item);
			object returnItem = Invoker.MethodReturn(this, "DDERequest", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837904.aspx
		/// </summary>
		/// <param name="channel">Int32 Channel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DDETerminate(Int32 channel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(channel);
			Invoker.Method(this, "DDETerminate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192053.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DDETerminateAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DDETerminateAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		/// <param name="arg4">optional object Arg4</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3, object arg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3, arg4);
			object returnItem = Invoker.MethodReturn(this, "BuildKeyCode", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey Arg1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1);
			object returnItem = Invoker.MethodReturn(this, "BuildKeyCode", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "BuildKeyCode", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx
		/// </summary>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey Arg1</param>
		/// <param name="arg2">optional object Arg2</param>
		/// <param name="arg3">optional object Arg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(arg1, arg2, arg3);
			object returnItem = Invoker.MethodReturn(this, "BuildKeyCode", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx
		/// </summary>
		/// <param name="keyCode">Int32 KeyCode</param>
		/// <param name="keyCode2">optional object KeyCode2</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string KeyString(Int32 keyCode, object keyCode2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(keyCode, keyCode2);
			object returnItem = Invoker.MethodReturn(this, "KeyString", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx
		/// </summary>
		/// <param name="keyCode">Int32 KeyCode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string KeyString(Int32 keyCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(keyCode);
			object returnItem = Invoker.MethodReturn(this, "KeyString", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835492.aspx
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="destination">string Destination</param>
		/// <param name="name">string Name</param>
		/// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject Object</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OrganizerCopy(string source, string destination, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, destination, name, _object);
			Invoker.Method(this, "OrganizerCopy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194744.aspx
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="name">string Name</param>
		/// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject Object</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OrganizerDelete(string source, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, name, _object);
			Invoker.Method(this, "OrganizerDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836140.aspx
		/// </summary>
		/// <param name="source">string Source</param>
		/// <param name="name">string Name</param>
		/// <param name="newName">string NewName</param>
		/// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject Object</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OrganizerRename(string source, string name, string newName, NetOffice.WordApi.Enums.WdOrganizerObject _object)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, name, newName, _object);
			Invoker.Method(this, "OrganizerRename", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823266.aspx
		/// </summary>
		/// <param name="tagID">String[] TagID</param>
		/// <param name="value">String[] Value</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AddAddress(String[] tagID, String[] value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)tagID, (object)value);
			Invoker.Method(this, "AddAddress", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		/// <param name="useAutoText">optional object UseAutoText</param>
		/// <param name="displaySelectDialog">optional object DisplaySelectDialog</param>
		/// <param name="selectDialog">optional object SelectDialog</param>
		/// <param name="checkNamesDialog">optional object CheckNamesDialog</param>
		/// <param name="recentAddressesChoice">optional object RecentAddressesChoice</param>
		/// <param name="updateRecentAddresses">optional object UpdateRecentAddresses</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice, object updateRecentAddresses)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice, updateRecentAddresses);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		/// <param name="useAutoText">optional object UseAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties, useAutoText);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		/// <param name="useAutoText">optional object UseAutoText</param>
		/// <param name="displaySelectDialog">optional object DisplaySelectDialog</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties, useAutoText, displaySelectDialog);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		/// <param name="useAutoText">optional object UseAutoText</param>
		/// <param name="displaySelectDialog">optional object DisplaySelectDialog</param>
		/// <param name="selectDialog">optional object SelectDialog</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties, useAutoText, displaySelectDialog, selectDialog);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		/// <param name="useAutoText">optional object UseAutoText</param>
		/// <param name="displaySelectDialog">optional object DisplaySelectDialog</param>
		/// <param name="selectDialog">optional object SelectDialog</param>
		/// <param name="checkNamesDialog">optional object CheckNamesDialog</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="addressProperties">optional object AddressProperties</param>
		/// <param name="useAutoText">optional object UseAutoText</param>
		/// <param name="displaySelectDialog">optional object DisplaySelectDialog</param>
		/// <param name="selectDialog">optional object SelectDialog</param>
		/// <param name="checkNamesDialog">optional object CheckNamesDialog</param>
		/// <param name="recentAddressesChoice">optional object RecentAddressesChoice</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice);
			object returnItem = Invoker.MethodReturn(this, "GetAddress", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194798.aspx
		/// </summary>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckGrammar(string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(this, "CheckGrammar", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		/// <param name="customDictionary10">optional object CustomDictionary10</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9);
			object returnItem = Invoker.MethodReturn(this, "CheckSpelling", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822681.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ResetIgnoreAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetIgnoreAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		/// <param name="customDictionary10">optional object CustomDictionary10</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx
		/// </summary>
		/// <param name="word">string Word</param>
		/// <param name="customDictionary">optional object CustomDictionary</param>
		/// <param name="ignoreUppercase">optional object IgnoreUppercase</param>
		/// <param name="mainDictionary">optional object MainDictionary</param>
		/// <param name="suggestionMode">optional object SuggestionMode</param>
		/// <param name="customDictionary2">optional object CustomDictionary2</param>
		/// <param name="customDictionary3">optional object CustomDictionary3</param>
		/// <param name="customDictionary4">optional object CustomDictionary4</param>
		/// <param name="customDictionary5">optional object CustomDictionary5</param>
		/// <param name="customDictionary6">optional object CustomDictionary6</param>
		/// <param name="customDictionary7">optional object CustomDictionary7</param>
		/// <param name="customDictionary8">optional object CustomDictionary8</param>
		/// <param name="customDictionary9">optional object CustomDictionary9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9);
			object returnItem = Invoker.MethodReturn(this, "GetSpellingSuggestions", paramsArray);
			NetOffice.WordApi.SpellingSuggestions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType) as NetOffice.WordApi.SpellingSuggestions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838545.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void GoBack()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "GoBack", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841057.aspx
		/// </summary>
		/// <param name="helpType">object HelpType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Help(object helpType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(helpType);
			Invoker.Method(this, "Help", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194337.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutomaticChange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutomaticChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839095.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ShowMe()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowMe", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821932.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void HelpTool()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "HelpTool", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845336.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194509.aspx
		/// </summary>
		/// <param name="listAllCommands">bool ListAllCommands</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ListCommands(bool listAllCommands)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(listAllCommands);
			Invoker.Method(this, "ListCommands", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834517.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ShowClipboard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowClipboard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx
		/// </summary>
		/// <param name="when">object When</param>
		/// <param name="name">string Name</param>
		/// <param name="tolerance">optional object Tolerance</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OnTime(object when, string name, object tolerance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, name, tolerance);
			Invoker.Method(this, "OnTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx
		/// </summary>
		/// <param name="when">object When</param>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void OnTime(object when, string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(when, name);
			Invoker.Method(this, "OnTime", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837154.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void NextLetter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "NextLetter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="zone">string Zone</param>
		/// <param name="server">string Server</param>
		/// <param name="volume">string Volume</param>
		/// <param name="user">optional object User</param>
		/// <param name="userPassword">optional object UserPassword</param>
		/// <param name="volumePassword">optional object VolumePassword</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume, object user, object userPassword, object volumePassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(zone, server, volume, user, userPassword, volumePassword);
			object returnItem = Invoker.MethodReturn(this, "MountVolume", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="zone">string Zone</param>
		/// <param name="server">string Server</param>
		/// <param name="volume">string Volume</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(zone, server, volume);
			object returnItem = Invoker.MethodReturn(this, "MountVolume", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="zone">string Zone</param>
		/// <param name="server">string Server</param>
		/// <param name="volume">string Volume</param>
		/// <param name="user">optional object User</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume, object user)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(zone, server, volume, user);
			object returnItem = Invoker.MethodReturn(this, "MountVolume", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="zone">string Zone</param>
		/// <param name="server">string Server</param>
		/// <param name="volume">string Volume</param>
		/// <param name="user">optional object User</param>
		/// <param name="userPassword">optional object UserPassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume, object user, object userPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(zone, server, volume, user, userPassword);
			object returnItem = Invoker.MethodReturn(this, "MountVolume", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844818.aspx
		/// </summary>
		/// <param name="_string">string String</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string CleanString(string _string)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(_string);
			object returnItem = Invoker.MethodReturn(this, "CleanString", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SendFax()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendFax", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835219.aspx
		/// </summary>
		/// <param name="path">string Path</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ChangeFileOpenDirectory(string path)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(path);
			Invoker.Method(this, "ChangeFileOpenDirectory", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void RunOld(string macroName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName);
			Invoker.Method(this, "RunOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196922.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void GoForward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "GoForward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844914.aspx
		/// </summary>
		/// <param name="left">Int32 Left</param>
		/// <param name="top">Int32 Top</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Move(Int32 left, Int32 top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197452.aspx
		/// </summary>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Resize(Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(width, height);
			Invoker.Method(this, "Resize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197549.aspx
		/// </summary>
		/// <param name="inches">Single Inches</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single InchesToPoints(Single inches)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(inches);
			object returnItem = Invoker.MethodReturn(this, "InchesToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838159.aspx
		/// </summary>
		/// <param name="centimeters">Single Centimeters</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single CentimetersToPoints(Single centimeters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(centimeters);
			object returnItem = Invoker.MethodReturn(this, "CentimetersToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845767.aspx
		/// </summary>
		/// <param name="millimeters">Single Millimeters</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single MillimetersToPoints(Single millimeters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(millimeters);
			object returnItem = Invoker.MethodReturn(this, "MillimetersToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840225.aspx
		/// </summary>
		/// <param name="picas">Single Picas</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PicasToPoints(Single picas)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(picas);
			object returnItem = Invoker.MethodReturn(this, "PicasToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840343.aspx
		/// </summary>
		/// <param name="lines">Single Lines</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single LinesToPoints(Single lines)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lines);
			object returnItem = Invoker.MethodReturn(this, "LinesToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838268.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToInches(Single points)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points);
			object returnItem = Invoker.MethodReturn(this, "PointsToInches", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195052.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToCentimeters(Single points)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points);
			object returnItem = Invoker.MethodReturn(this, "PointsToCentimeters", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836929.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToMillimeters(Single points)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points);
			object returnItem = Invoker.MethodReturn(this, "PointsToMillimeters", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193434.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToPicas(Single points)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points);
			object returnItem = Invoker.MethodReturn(this, "PointsToPicas", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822110.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToLines(Single points)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points);
			object returnItem = Invoker.MethodReturn(this, "PointsToLines", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821351.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		/// <param name="fVertical">optional object fVertical</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToPixels(Single points, object fVertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points, fVertical);
			object returnItem = Invoker.MethodReturn(this, "PointsToPixels", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx
		/// </summary>
		/// <param name="points">Single Points</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PointsToPixels(Single points)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(points);
			object returnItem = Invoker.MethodReturn(this, "PointsToPixels", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx
		/// </summary>
		/// <param name="pixels">Single Pixels</param>
		/// <param name="fVertical">optional object fVertical</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PixelsToPoints(Single pixels, object fVertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pixels, fVertical);
			object returnItem = Invoker.MethodReturn(this, "PixelsToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx
		/// </summary>
		/// <param name="pixels">Single Pixels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PixelsToPoints(Single pixels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pixels);
			object returnItem = Invoker.MethodReturn(this, "PixelsToPoints", paramsArray);
			return NetRuntimeSystem.Convert.ToSingle(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845662.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void KeyboardLatin()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "KeyboardLatin", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196621.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void KeyboardBidi()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "KeyboardBidi", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835971.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void ToggleKeyboard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ToggleKeyboard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx
		/// </summary>
		/// <param name="langId">optional Int32 LangId = 0</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Keyboard(object langId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(langId);
			object returnItem = Invoker.MethodReturn(this, "Keyboard", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Keyboard()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Keyboard", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193728.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ProductCode()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ProductCode", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840160.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.DefaultWebOptions DefaultWebOptions()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DefaultWebOptions", paramsArray);
			NetOffice.WordApi.DefaultWebOptions newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.DefaultWebOptions.LateBindingApiWrapperType) as NetOffice.WordApi.DefaultWebOptions;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">object Range</param>
		/// <param name="cid">object cid</param>
		/// <param name="piCSE">object piCSE</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void DiscussionSupport(object range, object cid, object piCSE)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, cid, piCSE);
			Invoker.Method(this, "DiscussionSupport", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821531.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium DocumentType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SetDefaultTheme(string name, NetOffice.WordApi.Enums.WdDocumentMedium documentType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, documentType);
			Invoker.Method(this, "SetDefaultTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834585.aspx
		/// </summary>
		/// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium DocumentType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string GetDefaultTheme(NetOffice.WordApi.Enums.WdDocumentMedium documentType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(documentType);
			object returnItem = Invoker.MethodReturn(this, "GetDefaultTheme", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object PrintZoomPaperHeight</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		/// <param name="varg25">optional object varg25</param>
		/// <param name="varg26">optional object varg26</param>
		/// <param name="varg27">optional object varg27</param>
		/// <param name="varg28">optional object varg28</param>
		/// <param name="varg29">optional object varg29</param>
		/// <param name="varg30">optional object varg30</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29, object varg30)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		/// <param name="varg25">optional object varg25</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		/// <param name="varg25">optional object varg25</param>
		/// <param name="varg26">optional object varg26</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		/// <param name="varg25">optional object varg25</param>
		/// <param name="varg26">optional object varg26</param>
		/// <param name="varg27">optional object varg27</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		/// <param name="varg25">optional object varg25</param>
		/// <param name="varg26">optional object varg26</param>
		/// <param name="varg27">optional object varg27</param>
		/// <param name="varg28">optional object varg28</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx
		/// </summary>
		/// <param name="macroName">string MacroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		/// <param name="varg10">optional object varg10</param>
		/// <param name="varg11">optional object varg11</param>
		/// <param name="varg12">optional object varg12</param>
		/// <param name="varg13">optional object varg13</param>
		/// <param name="varg14">optional object varg14</param>
		/// <param name="varg15">optional object varg15</param>
		/// <param name="varg16">optional object varg16</param>
		/// <param name="varg17">optional object varg17</param>
		/// <param name="varg18">optional object varg18</param>
		/// <param name="varg19">optional object varg19</param>
		/// <param name="varg20">optional object varg20</param>
		/// <param name="varg21">optional object varg21</param>
		/// <param name="varg22">optional object varg22</param>
		/// <param name="varg23">optional object varg23</param>
		/// <param name="varg24">optional object varg24</param>
		/// <param name="varg25">optional object varg25</param>
		/// <param name="varg26">optional object varg26</param>
		/// <param name="varg27">optional object varg27</param>
		/// <param name="varg28">optional object varg28</param>
		/// <param name="varg29">optional object varg29</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29);
			object returnItem = Invoker.MethodReturn(this, "Run", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object PrintZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight);
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
		/// <param name="fileName">optional object FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow);
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
		/// <param name="fileName">optional object FileName</param>
		/// <param name="activePrinterMacGX">optional object ActivePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object ManualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object PrintZoomColumn</param>
		/// <param name="printZoomRow">optional object PrintZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object PrintZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth);
			Invoker.Method(this, "PrintOut2000", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool Dummy2()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Dummy2", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838158.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14,15,16)]
		public void PutFocusInMailHeader()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PutFocusInMailHeader", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840673.aspx
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void LoadMasterList(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "LoadMasterList", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		/// <param name="revisedAuthor">optional string RevisedAuthor = </param>
		/// <param name="ignoreAllComparisonWarnings">optional bool IgnoreAllComparisonWarnings = false</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor, object ignoreAllComparisonWarnings)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor, ignoreAllComparisonWarnings);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		/// <param name="revisedAuthor">optional string RevisedAuthor = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor);
			object returnItem = Invoker.MethodReturn(this, "CompareDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		/// <param name="originalAuthor">optional string OriginalAuthor = </param>
		/// <param name="revisedAuthor">optional string RevisedAuthor = </param>
		/// <param name="formatFrom">optional NetOffice.WordApi.Enums.WdMergeFormatFrom FormatFrom = 2</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor, object formatFrom)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor, formatFrom);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		/// <param name="originalAuthor">optional string OriginalAuthor = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx
		/// </summary>
		/// <param name="originalDocument">NetOffice.WordApi.Document OriginalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document RevisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		/// <param name="compareFields">optional bool CompareFields = true</param>
		/// <param name="compareComments">optional bool CompareComments = true</param>
		/// <param name="compareMoves">optional bool CompareMoves = true</param>
		/// <param name="originalAuthor">optional string OriginalAuthor = </param>
		/// <param name="revisedAuthor">optional string RevisedAuthor = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor);
			object returnItem = Invoker.MethodReturn(this, "MergeDocuments", paramsArray);
			NetOffice.WordApi.Document newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Document.LateBindingApiWrapperType) as NetOffice.WordApi.Document;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="localDocument">NetOffice.WordApi.Document LocalDocument</param>
		/// <param name="serverDocument">NetOffice.WordApi.Document ServerDocument</param>
		/// <param name="baseDocument">NetOffice.WordApi.Document BaseDocument</param>
		/// <param name="favorSource">bool FavorSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void ThreeWayMerge(NetOffice.WordApi.Document localDocument, NetOffice.WordApi.Document serverDocument, NetOffice.WordApi.Document baseDocument, bool favorSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(localDocument, serverDocument, baseDocument, favorSource);
			Invoker.Method(this, "ThreeWayMerge", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public void Dummy4()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Dummy4", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}