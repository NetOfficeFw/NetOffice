using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{  
    /// <summary>
    /// DispatchInterface _Application 
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Presentations"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentations Presentations
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Presentations>(this, "Presentations", NetOffice.PowerPointApi.Presentations.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Windows"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DocumentWindows>(this, "Windows", NetOffice.PowerPointApi.DocumentWindows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.PPDialogs Dialogs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PPDialogs>(this, "Dialogs", NetOffice.PowerPointApi.PPDialogs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ActiveWindow"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindow ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DocumentWindow>(this, "ActiveWindow", NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ActivePresentation"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Presentation ActivePresentation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Presentation>(this, "ActivePresentation", NetOffice.PowerPointApi.Presentation.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SlideShowWindows"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowWindows SlideShowWindows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowWindows>(this, "SlideShowWindows", NetOffice.PowerPointApi.SlideShowWindows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.CommandBars"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Path"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Name"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Caption"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileSearch FileSearch
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(this, "FileSearch", NetOffice.OfficeApi.FileSearch.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.IFind FileFind
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IFind>(this, "FileFind", NetOffice.OfficeApi.IFind.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Build"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Build
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Version"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.OperatingSystem"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string OperatingSystem
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OperatingSystem");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ActivePrinter"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string ActivePrinter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ActivePrinter");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Creator"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AddIns"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.AddIns AddIns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.AddIns>(this, "AddIns", NetOffice.PowerPointApi.AddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.VBE"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Left"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Top"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Width"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Height"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.WindowState"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Visible"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Visible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Visible", value);
			}
		}

        /// <summary>
        /// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        [KnownIssue]
        [Obsolete("Int32 HWND is unavailable. Use ApplicationUtils instead.")]
        public virtual Int32 HWND
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Active"/> </remarks>
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
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.AnswerWizard AnswerWizard
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(this, "AnswerWizard", NetOffice.OfficeApi.AnswerWizard.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.COMAddIns"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.COMAddIns COMAddIns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", NetOffice.OfficeApi.COMAddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProductCode"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string ProductCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProductCode");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DefaultWebOptions DefaultWebOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DefaultWebOptions>(this, "DefaultWebOptions", NetOffice.PowerPointApi.DefaultWebOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.LanguageSettings"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.LanguageSettings LanguageSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(this, "LanguageSettings", NetOffice.OfficeApi.LanguageSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.MsoDebugOptions MsoDebugOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions>(this, "MsoDebugOptions", NetOffice.OfficeApi.MsoDebugOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ShowWindowsInTaskbar"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ShowWindowsInTaskbar
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ShowWindowsInTaskbar");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowWindowsInTaskbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.Marker Marker
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Marker>(this, "Marker", NetOffice.PowerPointApi.Marker.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.FeatureInstall"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.FileDialog"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(this, "FileDialog", NetOffice.OfficeApi.FileDialog.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Alias for get_FileDialog
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.FileDialog"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoFileDialogType type</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16), Redirect("get_FileDialog")]
		public NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType type)
		{
			return get_FileDialog(type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.DisplayGridLines"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplayGridLines
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "DisplayGridLines");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayGridLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AutomationSecurity"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.NewPresentation(property)"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.NewFile NewPresentation
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(this, "NewPresentation", NetOffice.OfficeApi.NewFile.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.DisplayAlerts"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpAlertLevel DisplayAlerts
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpAlertLevel>(this, "DisplayAlerts");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayAlerts", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ShowStartupDialog"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ShowStartupDialog
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ShowStartupDialog");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowStartupDialog", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.AutoCorrect"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.AutoCorrect AutoCorrect
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.AutoCorrect>(this, "AutoCorrect", NetOffice.PowerPointApi.AutoCorrect.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Options"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Options Options
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Options>(this, "Options", NetOffice.PowerPointApi.Options.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.DisplayDocumentInformationPanel"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Assistance"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.IAssistance Assistance
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", NetOffice.OfficeApi.IAssistance.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ActiveEncryptionSession"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Int32 ActiveEncryptionSession
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ActiveEncryptionSession");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.FileConverters"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.FileConverters FileConverters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.FileConverters>(this, "FileConverters", NetOffice.PowerPointApi.FileConverters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SmartArtLayouts"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtLayouts>(this, "SmartArtLayouts", NetOffice.OfficeApi.SmartArtLayouts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SmartArtQuickStyles"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtQuickStyles>(this, "SmartArtQuickStyles", NetOffice.OfficeApi.SmartArtQuickStyles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.SmartArtColors"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtColors SmartArtColors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtColors>(this, "SmartArtColors", NetOffice.OfficeApi.SmartArtColors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ProtectedViewWindows"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ProtectedViewWindows ProtectedViewWindows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ProtectedViewWindows>(this, "ProtectedViewWindows", NetOffice.PowerPointApi.ProtectedViewWindows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.ActiveProtectedViewWindow"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ProtectedViewWindow ActiveProtectedViewWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ProtectedViewWindow>(this, "ActiveProtectedViewWindow", NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.IsSandboxed"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool IsSandboxed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSandboxed");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ResampleMediaTasks ResampleMediaTasks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ResampleMediaTasks>(this, "ResampleMediaTasks", NetOffice.PowerPointApi.ResampleMediaTasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.FileValidation"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.application.chartdatapointtrack"/> </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
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
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.application.displayguides"/> </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplayGuides
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "DisplayGuides");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayGuides", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Help"/> </remarks>
		/// <param name="helpFile">optional string HelpFile = vbappt9.chm</param>
		/// <param name="contextID">optional Int32 ContextID = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Help(object helpFile, object contextID)
		{
			 Factory.ExecuteMethod(this, "Help", helpFile, contextID);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Help"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Help()
		{
			 Factory.ExecuteMethod(this, "Help");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Help"/> </remarks>
		/// <param name="helpFile">optional string HelpFile = vbappt9.chm</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Help(object helpFile)
		{
			 Factory.ExecuteMethod(this, "Help", helpFile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Quit"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Quit()
		{
			 Factory.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Run"/> </remarks>
		/// <param name="macroName">string macroName</param>
		/// <param name="safeArrayOfParams">optional object[] safeArrayOfParams</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public object Run(string macroName, object[] safeArrayOfParams)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", macroName, safeArrayOfParams);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Run"/> </remarks>
		/// <param name="macroName">string macroName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public object Run(string macroName)
		{
			return Factory.ExecuteVariantMethodGet(this, "Run", macroName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpFileDialogType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.FileDialog FileDialog(NetOffice.PowerPointApi.Enums.PpFileDialogType type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.FileDialog>(this, "FileDialog", NetOffice.PowerPointApi.FileDialog.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pWindow">NetOffice.PowerPointApi.DocumentWindow pWindow</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void LaunchSpelling(NetOffice.PowerPointApi.DocumentWindow pWindow)
		{
			 Factory.ExecuteMethod(this, "LaunchSpelling", pWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.Activate"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="persist">optional bool Persist = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public bool GetOptionFlag(Int32 option, object persist)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetOptionFlag", option, persist);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public bool GetOptionFlag(Int32 option)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetOptionFlag", option);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="state">bool state</param>
		/// <param name="persist">optional bool Persist = false</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SetOptionFlag(Int32 option, bool state, object persist)
		{
			 Factory.ExecuteMethod(this, "SetOptionFlag", option, state, persist);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="option">Int32 option</param>
		/// <param name="state">bool state</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SetOptionFlag(Int32 option, bool state)
		{
			 Factory.ExecuteMethod(this, "SetOptionFlag", option, state);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpFileDialogType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public object PPFileDialog(NetOffice.PowerPointApi.Enums.PpFileDialogType type)
		{
			return Factory.ExecuteVariantMethodGet(this, "PPFileDialog", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="marker">Int32 marker</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SetPerfMarker(Int32 marker)
		{
			 Factory.ExecuteMethod(this, "SetPerfMarker", marker);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void LaunchPublishSlidesDialog(string slideLibraryUrl)
		{
			 Factory.ExecuteMethod(this, "LaunchPublishSlidesDialog", slideLibraryUrl);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="slideUrls">object slideUrls</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void LaunchSendToPPTDialog(object slideUrls)
		{
			 Factory.ExecuteMethod(this, "LaunchSendToPPTDialog", slideUrls);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Application.StartNewUndoEntry"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void StartNewUndoEntry()
		{
			 Factory.ExecuteMethod(this, "StartNewUndoEntry");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.application.openthemefile"/> </remarks>
		/// <param name="themeFileName">string themeFileName</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Theme OpenThemeFile(string themeFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Theme>(this, "OpenThemeFile", NetOffice.PowerPointApi.Theme.LateBindingApiWrapperType, themeFileName);
		}

		#endregion

		#pragma warning restore
	}
}
