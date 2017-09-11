using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface _Application 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Application : COMObject
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
                    _type = typeof(_Application);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Application(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Application(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823254.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197825.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191758.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845178.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821628.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Documents Documents
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Documents>(this, "Documents", NetOffice.WordApi.Documents.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822351.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Windows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Windows>(this, "Windows", NetOffice.WordApi.Windows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837737.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document ActiveDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "ActiveDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845301.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(this, "ActiveWindow", NetOffice.WordApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838682.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Selection Selection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Selection>(this, "Selection", NetOffice.WordApi.Selection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822917.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object WordBasic
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "WordBasic");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195679.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.RecentFiles RecentFiles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RecentFiles>(this, "RecentFiles", NetOffice.WordApi.RecentFiles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845589.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Template NormalTemplate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Template>(this, "NormalTemplate", NetOffice.WordApi.Template.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822391.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.System System
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.System>(this, "System", NetOffice.WordApi.System.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845308.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCorrect AutoCorrect
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrect>(this, "AutoCorrect", NetOffice.WordApi.AutoCorrect.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197817.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FontNames FontNames
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(this, "FontNames", NetOffice.WordApi.FontNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196340.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FontNames LandscapeFontNames
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(this, "LandscapeFontNames", NetOffice.WordApi.FontNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192201.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FontNames PortraitFontNames
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(this, "PortraitFontNames", NetOffice.WordApi.FontNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840701.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Languages Languages
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Languages>(this, "Languages", NetOffice.WordApi.Languages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821300.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Browser Browser
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Browser>(this, "Browser", NetOffice.WordApi.Browser.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823259.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FileConverters FileConverters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FileConverters>(this, "FileConverters", NetOffice.WordApi.FileConverters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821659.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailingLabel MailingLabel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailingLabel>(this, "MailingLabel", NetOffice.WordApi.MailingLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191745.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Dialogs Dialogs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Dialogs>(this, "Dialogs", NetOffice.WordApi.Dialogs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838479.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.CaptionLabels CaptionLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CaptionLabels>(this, "CaptionLabels", NetOffice.WordApi.CaptionLabels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198063.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCaptions AutoCaptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCaptions>(this, "AutoCaptions", NetOffice.WordApi.AutoCaptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822986.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AddIns AddIns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AddIns>(this, "AddIns", NetOffice.WordApi.AddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839544.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821519.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197438.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ScreenUpdating
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ScreenUpdating");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScreenUpdating", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198164.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintPreview
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintPreview");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839740.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tasks Tasks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tasks>(this, "Tasks", NetOffice.WordApi.Tasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayStatusBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayStatusBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayStatusBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836086.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SpecialMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpecialMode");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839688.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 UsableWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UsableWidth");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834606.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 UsableHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UsableHeight");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192165.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MathCoprocessorAvailable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MathCoprocessorAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192426.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MouseAvailable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MouseAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
		/// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_International(NetOffice.WordApi.Enums.WdInternationalIndex index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "International", index);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_International
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
		/// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_International")]
		public object International(NetOffice.WordApi.Enums.WdInternationalIndex index)
		{
			return get_International(index);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839495.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Build
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820850.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CapsLock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CapsLock");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845392.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool NumLock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NumLock");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834599.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string UserName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844813.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string UserInitials
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserInitials");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserInitials", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193411.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string UserAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835128.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object MacroContainer
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "MacroContainer");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838964.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayRecentFiles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayRecentFiles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayRecentFiles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845623.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word, object languageID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(this, "SynonymInfo", NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType, word, languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_SynonymInfo
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_SynonymInfo")]
		public NetOffice.WordApi.SynonymInfo SynonymInfo(string word, object languageID)
		{
			return get_SynonymInfo(word, languageID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
		/// <param name="word">string word</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(this, "SynonymInfo", NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType, word);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_SynonymInfo
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
		/// <param name="word">string word</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_SynonymInfo")]
		public NetOffice.WordApi.SynonymInfo SynonymInfo(string word)
		{
			return get_SynonymInfo(word);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197234.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839412.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DefaultSaveFormat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultSaveFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultSaveFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821102.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListGalleries ListGalleries
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListGalleries>(this, "ListGalleries", NetOffice.WordApi.ListGalleries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821995.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ActivePrinter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActivePrinter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ActivePrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821925.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Templates Templates
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Templates>(this, "Templates", NetOffice.WordApi.Templates.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822548.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object CustomizationContext
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CustomizationContext");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "CustomizationContext", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197596.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.KeyBindings KeyBindings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBindings>(this, "KeyBindings", NetOffice.WordApi.KeyBindings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
		/// <param name="command">string command</param>
		/// <param name="commandParameter">optional object commandParameter</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeysBoundTo>(this, "KeysBoundTo", NetOffice.WordApi.KeysBoundTo.LateBindingApiWrapperType, keyCategory, command, commandParameter);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_KeysBoundTo
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
		/// <param name="command">string command</param>
		/// <param name="commandParameter">optional object commandParameter</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_KeysBoundTo")]
		public NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
		{
			return get_KeysBoundTo(keyCategory, command, commandParameter);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
		/// <param name="command">string command</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeysBoundTo>(this, "KeysBoundTo", NetOffice.WordApi.KeysBoundTo.LateBindingApiWrapperType, keyCategory, command);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_KeysBoundTo
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
		/// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
		/// <param name="command">string command</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_KeysBoundTo")]
		public NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
		{
			return get_KeysBoundTo(keyCategory, command);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
		/// <param name="keyCode">Int32 keyCode</param>
		/// <param name="keyCode2">optional object keyCode2</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode, object keyCode2)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBinding>(this, "FindKey", NetOffice.WordApi.KeyBinding.LateBindingApiWrapperType, keyCode, keyCode2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_FindKey
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
		/// <param name="keyCode">Int32 keyCode</param>
		/// <param name="keyCode2">optional object keyCode2</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_FindKey")]
		public NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode, object keyCode2)
		{
			return get_FindKey(keyCode, keyCode2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
		/// <param name="keyCode">Int32 keyCode</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBinding>(this, "FindKey", NetOffice.WordApi.KeyBinding.LateBindingApiWrapperType, keyCode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_FindKey
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
		/// <param name="keyCode">Int32 keyCode</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_FindKey")]
		public NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode)
		{
			return get_FindKey(keyCode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196028.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192216.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192367.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayScrollBars
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayScrollBars");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191937.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string StartupPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StartupPath");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartupPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835146.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BackgroundSavingStatus
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackgroundSavingStatus");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820962.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BackgroundPrintingStatus
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackgroundPrintingStatus");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839318.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837463.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836284.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845159.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836388.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdWindowState WindowState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdWindowState>(this, "WindowState");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192152.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayAutoCompleteTips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAutoCompleteTips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAutoCompleteTips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822542.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Options Options
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Options>(this, "Options", NetOffice.WordApi.Options.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192373.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdAlertLevel DisplayAlerts
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdAlertLevel>(this, "DisplayAlerts");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayAlerts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191957.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Dictionaries CustomDictionaries
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Dictionaries>(this, "CustomDictionaries", NetOffice.WordApi.Dictionaries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192616.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string PathSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PathSeparator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845291.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string StatusBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StatusBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StatusBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192800.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MAPIAvailable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MAPIAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845182.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayScreenTips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayScreenTips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayScreenTips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839294.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdEnableCancelKey EnableCancelKey
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEnableCancelKey>(this, "EnableCancelKey");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EnableCancelKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197424.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UserControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UserControl");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(this, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838972.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMailSystem MailSystem
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailSystem>(this, "MailSystem");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839937.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DefaultTableSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultTableSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultTableSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839922.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowVisualBasicEditor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowVisualBasicEditor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowVisualBasicEditor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839549.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string BrowseExtraFileTypes
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BrowseExtraFileTypes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BrowseExtraFileTypes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
		/// <param name="_object">object object</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsObjectValid(object _object)
		{
			return Factory.ExecuteBoolPropertyGet(this, "IsObjectValid", _object);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_IsObjectValid
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
		/// <param name="_object">object object</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_IsObjectValid")]
		public bool IsObjectValid(object _object)
		{
			return get_IsObjectValid(_object);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194713.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.HangulHanjaConversionDictionaries HangulHanjaDictionaries
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HangulHanjaConversionDictionaries>(this, "HangulHanjaDictionaries", NetOffice.WordApi.HangulHanjaConversionDictionaries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821986.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMessage MailMessage
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMessage>(this, "MailMessage", NetOffice.WordApi.MailMessage.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840871.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool FocusInMailHeader
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FocusInMailHeader");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192588.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.EmailOptions EmailOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.EmailOptions>(this, "EmailOptions", NetOffice.WordApi.EmailOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836711.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID Language
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "Language");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192831.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192428.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckLanguage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckLanguage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckLanguage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197161.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(this, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy1
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dummy1");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(this, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195192.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFeatureInstall>(this, "FeatureInstall");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FeatureInstall", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192776.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutomationSecurity>(this, "AutomationSecurity");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutomationSecurity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
		/// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(this, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, fileDialogType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
		/// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
		[SupportByVersion("Word", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		public NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
		{
			return get_FileDialog(fileDialogType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193382.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string EmailTemplate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EmailTemplate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmailTemplate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowWindowsInTaskbar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowWindowsInTaskbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowWindowsInTaskbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193065.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.NewFile NewDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(this, "NewDocument", NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840052.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowStartupDialog
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowStartupDialog");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowStartupDialog", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192177.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCorrect AutoCorrectEmail
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrect>(this, "AutoCorrectEmail", NetOffice.WordApi.AutoCorrect.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845341.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.TaskPanes TaskPanes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TaskPanes>(this, "TaskPanes", NetOffice.WordApi.TaskPanes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835491.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DefaultLegalBlackline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultLegalBlackline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultLegalBlackline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.SmartTagRecognizers SmartTagRecognizers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTagRecognizers>(this, "SmartTagRecognizers", NetOffice.WordApi.SmartTagRecognizers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.SmartTagTypes SmartTagTypes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTagTypes>(this, "SmartTagTypes", NetOffice.WordApi.SmartTagTypes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839771.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNamespaces XMLNamespaces
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNamespaces>(this, "XMLNamespaces", NetOffice.WordApi.XMLNamespaces.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196679.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ArbitraryXMLSupportAvailable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ArbitraryXMLSupportAvailable");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BuildFull
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildFull");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BuildFeatureCrew
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildFeatureCrew");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192405.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Bibliography Bibliography
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bibliography>(this, "Bibliography", NetOffice.WordApi.Bibliography.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191727.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowStylePreviews
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowStylePreviews");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowStylePreviews", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845435.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool RestrictLinkedStyles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RestrictLinkedStyles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RestrictLinkedStyles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837322.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMathAutoCorrect OMathAutoCorrect
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathAutoCorrect>(this, "OMathAutoCorrect", NetOffice.WordApi.OMathAutoCorrect.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836074.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool DisplayDocumentInformationPanel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayDocumentInformationPanel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayDocumentInformationPanel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197133.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192620.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool OpenAttachmentsInFullScreen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OpenAttachmentsInFullScreen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OpenAttachmentsInFullScreen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836063.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 ActiveEncryptionSession
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ActiveEncryptionSession");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194203.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool DontResetInsertionPointProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DontResetInsertionPointProperties");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DontResetInsertionPointProperties", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839192.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtLayouts>(this, "SmartArtLayouts", NetOffice.OfficeApi.SmartArtLayouts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194982.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtQuickStyles>(this, "SmartArtQuickStyles", NetOffice.OfficeApi.SmartArtQuickStyles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839505.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtColors SmartArtColors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtColors>(this, "SmartArtColors", NetOffice.OfficeApi.SmartArtColors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838675.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.UndoRecord UndoRecord
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.UndoRecord>(this, "UndoRecord", NetOffice.WordApi.UndoRecord.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191978.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.PickerDialog PickerDialog
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerDialog>(this, "PickerDialog", NetOffice.OfficeApi.PickerDialog.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839925.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindows ProtectedViewWindows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProtectedViewWindows>(this, "ProtectedViewWindows", NetOffice.WordApi.ProtectedViewWindows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192773.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow ActiveProtectedViewWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProtectedViewWindow>(this, "ActiveProtectedViewWindow", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845787.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool IsSandboxed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSandboxed");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193078.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileValidationMode>(this, "FileValidation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FileValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232091.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool ChartDataPointTrack
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ChartDataPointTrack");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ChartDataPointTrack", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232207.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool ShowAnimation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAnimation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAnimation", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Quit(object saveChanges, object originalFormat, object routeDocument)
		{
			 Factory.ExecuteMethod(this, "Quit", saveChanges, originalFormat, routeDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Quit(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "Quit", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Quit(object saveChanges, object originalFormat)
		{
			 Factory.ExecuteMethod(this, "Quit", saveChanges, originalFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193095.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ScreenRefresh()
		{
			 Factory.ExecuteMethod(this, "ScreenRefresh");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld()
		{
			 Factory.ExecuteMethod(this, "PrintOutOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
		{
			 Factory.ExecuteMethod(this, "PrintOutOld", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839803.aspx </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void LookupNameProperties(string name)
		{
			 Factory.ExecuteMethod(this, "LookupNameProperties", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192415.aspx </remarks>
		/// <param name="unavailableFont">string unavailableFont</param>
		/// <param name="substituteFont">string substituteFont</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SubstituteFont(string unavailableFont, string substituteFont)
		{
			 Factory.ExecuteMethod(this, "SubstituteFont", unavailableFont, substituteFont);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
		/// <param name="times">optional object times</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Repeat(object times)
		{
			return Factory.ExecuteBoolMethodGet(this, "Repeat", times);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Repeat()
		{
			return Factory.ExecuteBoolMethodGet(this, "Repeat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845561.aspx </remarks>
		/// <param name="channel">Int32 channel</param>
		/// <param name="command">string command</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DDEExecute(Int32 channel, string command)
		{
			 Factory.ExecuteMethod(this, "DDEExecute", channel, command);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837295.aspx </remarks>
		/// <param name="app">string app</param>
		/// <param name="topic">string topic</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 DDEInitiate(string app, string topic)
		{
			return Factory.ExecuteInt32MethodGet(this, "DDEInitiate", app, topic);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837201.aspx </remarks>
		/// <param name="channel">Int32 channel</param>
		/// <param name="item">string item</param>
		/// <param name="data">string data</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DDEPoke(Int32 channel, string item, string data)
		{
			 Factory.ExecuteMethod(this, "DDEPoke", channel, item, data);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837546.aspx </remarks>
		/// <param name="channel">Int32 channel</param>
		/// <param name="item">string item</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DDERequest(Int32 channel, string item)
		{
			return Factory.ExecuteStringMethodGet(this, "DDERequest", channel, item);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837904.aspx </remarks>
		/// <param name="channel">Int32 channel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DDETerminate(Int32 channel)
		{
			 Factory.ExecuteMethod(this, "DDETerminate", channel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192053.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DDETerminateAll()
		{
			 Factory.ExecuteMethod(this, "DDETerminateAll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3, object arg4)
		{
			return Factory.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1, arg2, arg3, arg4);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1)
		{
			return Factory.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2)
		{
			return Factory.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
		/// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3)
		{
			return Factory.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1, arg2, arg3);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
		/// <param name="keyCode">Int32 keyCode</param>
		/// <param name="keyCode2">optional object keyCode2</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string KeyString(Int32 keyCode, object keyCode2)
		{
			return Factory.ExecuteStringMethodGet(this, "KeyString", keyCode, keyCode2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
		/// <param name="keyCode">Int32 keyCode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string KeyString(Int32 keyCode)
		{
			return Factory.ExecuteStringMethodGet(this, "KeyString", keyCode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835492.aspx </remarks>
		/// <param name="source">string source</param>
		/// <param name="destination">string destination</param>
		/// <param name="name">string name</param>
		/// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OrganizerCopy(string source, string destination, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
		{
			 Factory.ExecuteMethod(this, "OrganizerCopy", source, destination, name, _object);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194744.aspx </remarks>
		/// <param name="source">string source</param>
		/// <param name="name">string name</param>
		/// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OrganizerDelete(string source, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
		{
			 Factory.ExecuteMethod(this, "OrganizerDelete", source, name, _object);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836140.aspx </remarks>
		/// <param name="source">string source</param>
		/// <param name="name">string name</param>
		/// <param name="newName">string newName</param>
		/// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OrganizerRename(string source, string name, string newName, NetOffice.WordApi.Enums.WdOrganizerObject _object)
		{
			 Factory.ExecuteMethod(this, "OrganizerRename", source, name, newName, _object);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823266.aspx </remarks>
		/// <param name="tagID">String[] tagID</param>
		/// <param name="value">String[] value</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AddAddress(String[] tagID, String[] value)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)tagID, (object)value);
            Invoker.Method(this, "AddAddress", paramsArray); ;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		/// <param name="useAutoText">optional object useAutoText</param>
		/// <param name="displaySelectDialog">optional object displaySelectDialog</param>
		/// <param name="selectDialog">optional object selectDialog</param>
		/// <param name="checkNamesDialog">optional object checkNamesDialog</param>
		/// <param name="recentAddressesChoice">optional object recentAddressesChoice</param>
		/// <param name="updateRecentAddresses">optional object updateRecentAddresses</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice, object updateRecentAddresses)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", new object[]{ name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice, updateRecentAddresses });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress()
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", name, addressProperties);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		/// <param name="useAutoText">optional object useAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", name, addressProperties, useAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		/// <param name="useAutoText">optional object useAutoText</param>
		/// <param name="displaySelectDialog">optional object displaySelectDialog</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", name, addressProperties, useAutoText, displaySelectDialog);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		/// <param name="useAutoText">optional object useAutoText</param>
		/// <param name="displaySelectDialog">optional object displaySelectDialog</param>
		/// <param name="selectDialog">optional object selectDialog</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", new object[]{ name, addressProperties, useAutoText, displaySelectDialog, selectDialog });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		/// <param name="useAutoText">optional object useAutoText</param>
		/// <param name="displaySelectDialog">optional object displaySelectDialog</param>
		/// <param name="selectDialog">optional object selectDialog</param>
		/// <param name="checkNamesDialog">optional object checkNamesDialog</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", new object[]{ name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="addressProperties">optional object addressProperties</param>
		/// <param name="useAutoText">optional object useAutoText</param>
		/// <param name="displaySelectDialog">optional object displaySelectDialog</param>
		/// <param name="selectDialog">optional object selectDialog</param>
		/// <param name="checkNamesDialog">optional object checkNamesDialog</param>
		/// <param name="recentAddressesChoice">optional object recentAddressesChoice</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice)
		{
			return Factory.ExecuteStringMethodGet(this, "GetAddress", new object[]{ name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194798.aspx </remarks>
		/// <param name="_string">string string</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckGrammar(string _string)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckGrammar", _string);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		/// <param name="customDictionary10">optional object customDictionary10</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", word);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", word, customDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", word, customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", word, customDictionary, ignoreUppercase, mainDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			return Factory.ExecuteBoolMethodGet(this, "CheckSpelling", new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822681.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ResetIgnoreAll()
		{
			 Factory.ExecuteMethod(this, "ResetIgnoreAll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		/// <param name="customDictionary10">optional object customDictionary10</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word, customDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word, customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, word, customDictionary, ignoreUppercase, mainDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
		/// <param name="word">string word</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838545.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void GoBack()
		{
			 Factory.ExecuteMethod(this, "GoBack");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841057.aspx </remarks>
		/// <param name="helpType">object helpType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Help(object helpType)
		{
			 Factory.ExecuteMethod(this, "Help", helpType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194337.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutomaticChange()
		{
			 Factory.ExecuteMethod(this, "AutomaticChange");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839095.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ShowMe()
		{
			 Factory.ExecuteMethod(this, "ShowMe");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821932.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void HelpTool()
		{
			 Factory.ExecuteMethod(this, "HelpTool");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845336.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Window NewWindow()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Window>(this, "NewWindow", NetOffice.WordApi.Window.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194509.aspx </remarks>
		/// <param name="listAllCommands">bool listAllCommands</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ListCommands(bool listAllCommands)
		{
			 Factory.ExecuteMethod(this, "ListCommands", listAllCommands);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834517.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ShowClipboard()
		{
			 Factory.ExecuteMethod(this, "ShowClipboard");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
		/// <param name="when">object when</param>
		/// <param name="name">string name</param>
		/// <param name="tolerance">optional object tolerance</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OnTime(object when, string name, object tolerance)
		{
			 Factory.ExecuteMethod(this, "OnTime", when, name, tolerance);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
		/// <param name="when">object when</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OnTime(object when, string name)
		{
			 Factory.ExecuteMethod(this, "OnTime", when, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837154.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void NextLetter()
		{
			 Factory.ExecuteMethod(this, "NextLetter");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="zone">string zone</param>
		/// <param name="server">string server</param>
		/// <param name="volume">string volume</param>
		/// <param name="user">optional object user</param>
		/// <param name="userPassword">optional object userPassword</param>
		/// <param name="volumePassword">optional object volumePassword</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume, object user, object userPassword, object volumePassword)
		{
			return Factory.ExecuteInt16MethodGet(this, "MountVolume", new object[]{ zone, server, volume, user, userPassword, volumePassword });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="zone">string zone</param>
		/// <param name="server">string server</param>
		/// <param name="volume">string volume</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume)
		{
			return Factory.ExecuteInt16MethodGet(this, "MountVolume", zone, server, volume);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="zone">string zone</param>
		/// <param name="server">string server</param>
		/// <param name="volume">string volume</param>
		/// <param name="user">optional object user</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume, object user)
		{
			return Factory.ExecuteInt16MethodGet(this, "MountVolume", zone, server, volume, user);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="zone">string zone</param>
		/// <param name="server">string server</param>
		/// <param name="volume">string volume</param>
		/// <param name="user">optional object user</param>
		/// <param name="userPassword">optional object userPassword</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int16 MountVolume(string zone, string server, string volume, object user, object userPassword)
		{
			return Factory.ExecuteInt16MethodGet(this, "MountVolume", new object[]{ zone, server, volume, user, userPassword });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844818.aspx </remarks>
		/// <param name="_string">string string</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string CleanString(string _string)
		{
			return Factory.ExecuteStringMethodGet(this, "CleanString", _string);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SendFax()
		{
			 Factory.ExecuteMethod(this, "SendFax");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835219.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ChangeFileOpenDirectory(string path)
		{
			 Factory.ExecuteMethod(this, "ChangeFileOpenDirectory", path);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macroName">string macroName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RunOld(string macroName)
		{
			 Factory.ExecuteMethod(this, "RunOld", macroName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196922.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void GoForward()
		{
			 Factory.ExecuteMethod(this, "GoForward");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844914.aspx </remarks>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Move(Int32 left, Int32 top)
		{
			 Factory.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197452.aspx </remarks>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Resize(Int32 width, Int32 height)
		{
			 Factory.ExecuteMethod(this, "Resize", width, height);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197549.aspx </remarks>
		/// <param name="inches">Single inches</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single InchesToPoints(Single inches)
		{
			return Factory.ExecuteSingleMethodGet(this, "InchesToPoints", inches);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838159.aspx </remarks>
		/// <param name="centimeters">Single centimeters</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single CentimetersToPoints(Single centimeters)
		{
			return Factory.ExecuteSingleMethodGet(this, "CentimetersToPoints", centimeters);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845767.aspx </remarks>
		/// <param name="millimeters">Single millimeters</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single MillimetersToPoints(Single millimeters)
		{
			return Factory.ExecuteSingleMethodGet(this, "MillimetersToPoints", millimeters);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840225.aspx </remarks>
		/// <param name="picas">Single picas</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PicasToPoints(Single picas)
		{
			return Factory.ExecuteSingleMethodGet(this, "PicasToPoints", picas);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840343.aspx </remarks>
		/// <param name="lines">Single lines</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LinesToPoints(Single lines)
		{
			return Factory.ExecuteSingleMethodGet(this, "LinesToPoints", lines);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838268.aspx </remarks>
		/// <param name="points">Single points</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToInches(Single points)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToInches", points);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195052.aspx </remarks>
		/// <param name="points">Single points</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToCentimeters(Single points)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToCentimeters", points);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836929.aspx </remarks>
		/// <param name="points">Single points</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToMillimeters(Single points)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToMillimeters", points);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193434.aspx </remarks>
		/// <param name="points">Single points</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToPicas(Single points)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToPicas", points);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822110.aspx </remarks>
		/// <param name="points">Single points</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToLines(Single points)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToLines", points);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821351.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
		/// <param name="points">Single points</param>
		/// <param name="fVertical">optional object fVertical</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToPixels(Single points, object fVertical)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToPixels", points, fVertical);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
		/// <param name="points">Single points</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PointsToPixels(Single points)
		{
			return Factory.ExecuteSingleMethodGet(this, "PointsToPixels", points);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
		/// <param name="pixels">Single pixels</param>
		/// <param name="fVertical">optional object fVertical</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PixelsToPoints(Single pixels, object fVertical)
		{
			return Factory.ExecuteSingleMethodGet(this, "PixelsToPoints", pixels, fVertical);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
		/// <param name="pixels">Single pixels</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PixelsToPoints(Single pixels)
		{
			return Factory.ExecuteSingleMethodGet(this, "PixelsToPoints", pixels);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845662.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void KeyboardLatin()
		{
			 Factory.ExecuteMethod(this, "KeyboardLatin");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196621.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void KeyboardBidi()
		{
			 Factory.ExecuteMethod(this, "KeyboardBidi");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835971.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ToggleKeyboard()
		{
			 Factory.ExecuteMethod(this, "ToggleKeyboard");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
		/// <param name="langId">optional Int32 LangId = 0</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Keyboard(object langId)
		{
			return Factory.ExecuteInt32MethodGet(this, "Keyboard", langId);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Keyboard()
		{
			return Factory.ExecuteInt32MethodGet(this, "Keyboard");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193728.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ProductCode()
		{
			return Factory.ExecuteStringMethodGet(this, "ProductCode");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840160.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.DefaultWebOptions DefaultWebOptions()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DefaultWebOptions>(this, "DefaultWebOptions", NetOffice.WordApi.DefaultWebOptions.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">object range</param>
		/// <param name="cid">object cid</param>
		/// <param name="piCSE">object piCSE</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DiscussionSupport(object range, object cid, object piCSE)
		{
			 Factory.ExecuteMethod(this, "DiscussionSupport", range, cid, piCSE);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821531.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetDefaultTheme(string name, NetOffice.WordApi.Enums.WdDocumentMedium documentType)
		{
			 Factory.ExecuteMethod(this, "SetDefaultTheme", name, documentType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834585.aspx </remarks>
		/// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string GetDefaultTheme(NetOffice.WordApi.Enums.WdDocumentMedium documentType)
		{
			return Factory.ExecuteStringMethodGet(this, "GetDefaultTheme", documentType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29, object varg30)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", macroName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", macroName, varg1);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", macroName, varg1, varg2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", macroName, varg1, varg2, varg3);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="varg1">optional object varg1</param>
		/// <param name="varg2">optional object varg2</param>
		/// <param name="varg3">optional object varg3</param>
		/// <param name="varg4">optional object varg4</param>
		/// <param name="varg5">optional object varg5</param>
		/// <param name="varg6">optional object varg6</param>
		/// <param name="varg7">optional object varg7</param>
		/// <param name="varg8">optional object varg8</param>
		/// <param name="varg9">optional object varg9</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
		/// <param name="macroName">string macroName</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", new object[]{ macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		/// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000()
		{
			 Factory.ExecuteMethod(this, "PrintOut2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background, append);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background, append, range);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", background, append, range, outputFileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="background">optional object background</param>
		/// <param name="append">optional object append</param>
		/// <param name="range">optional object range</param>
		/// <param name="outputFileName">optional object outputFileName</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="item">optional object item</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="pages">optional object pages</param>
		/// <param name="pageType">optional object pageType</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
		/// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
		/// <param name="printZoomColumn">optional object printZoomColumn</param>
		/// <param name="printZoomRow">optional object printZoomRow</param>
		/// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
		{
			 Factory.ExecuteMethod(this, "PrintOut2000", new object[]{ background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool Dummy2()
		{
			return Factory.ExecuteBoolMethodGet(this, "Dummy2");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838158.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void PutFocusInMailHeader()
		{
			 Factory.ExecuteMethod(this, "PutFocusInMailHeader");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840673.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void LoadMasterList(string fileName)
		{
			 Factory.ExecuteMethod(this, "LoadMasterList", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor, object ignoreAllComparisonWarnings)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor, ignoreAllComparisonWarnings });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination, granularity);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor, object formatFrom)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor, formatFrom });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, originalDocument, revisedDocument, destination, granularity);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
		/// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
		/// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
		/// <param name="compareFormatting">optional bool CompareFormatting = true</param>
		/// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
		/// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
		/// <param name="compareTables">optional bool CompareTables = true</param>
		/// <param name="compareHeaders">optional bool CompareHeaders = true</param>
		/// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
		/// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
		/// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
		/// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
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
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="localDocument">NetOffice.WordApi.Document localDocument</param>
		/// <param name="serverDocument">NetOffice.WordApi.Document serverDocument</param>
		/// <param name="baseDocument">NetOffice.WordApi.Document baseDocument</param>
		/// <param name="favorSource">bool favorSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public void ThreeWayMerge(NetOffice.WordApi.Document localDocument, NetOffice.WordApi.Document serverDocument, NetOffice.WordApi.Document baseDocument, bool favorSource)
		{
			 Factory.ExecuteMethod(this, "ThreeWayMerge", localDocument, serverDocument, baseDocument, favorSource);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		public void Dummy4()
		{
			 Factory.ExecuteMethod(this, "Dummy4");
		}

		#endregion

		#pragma warning restore
	}
}
