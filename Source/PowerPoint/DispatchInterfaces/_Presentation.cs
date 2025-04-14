using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface _Presentation 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Presentation : COMObject
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
                    _type = typeof(_Presentation);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Presentation(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Presentation(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SlideMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master SlideMaster
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "SlideMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.TitleMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master TitleMaster
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "TitleMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.HasTitleMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasTitleMaster
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasTitleMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.TemplateName"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string TemplateName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TemplateName");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.NotesMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master NotesMaster
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "NotesMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.HandoutMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master HandoutMaster
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "HandoutMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Slides"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slides Slides
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Slides>(this, "Slides", NetOffice.PowerPointApi.Slides.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PageSetup"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PageSetup PageSetup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PageSetup>(this, "PageSetup", NetOffice.PowerPointApi.PageSetup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ColorSchemes"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorSchemes ColorSchemes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ColorSchemes>(this, "ColorSchemes", NetOffice.PowerPointApi.ColorSchemes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExtraColors"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ExtraColors ExtraColors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ExtraColors>(this, "ExtraColors", NetOffice.PowerPointApi.ExtraColors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SlideShowSettings"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowSettings SlideShowSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowSettings>(this, "SlideShowSettings", NetOffice.PowerPointApi.SlideShowSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Fonts"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Fonts Fonts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Fonts>(this, "Fonts", NetOffice.PowerPointApi.Fonts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Windows"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Tags"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Tags Tags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Tags>(this, "Tags", NetOffice.PowerPointApi.Tags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.DefaultShape"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape DefaultShape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "DefaultShape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.BuiltInDocumentProperties"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object BuiltInDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "BuiltInDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CustomDocumentProperties"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object CustomDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.VBProject"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ReadOnly"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ReadOnly
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ReadOnly");
			}
		}

		/// <summary>
		/// When you open a presentation that was saved as read-only recommended, Microsoft PowerPoint displays a message recommending that you open the presentation as read-only.
		///
		/// Use the SaveCopyAs2 method to change this property.
		///
		/// SupportByVersion PowerPoint 16
		/// </summary>
		/// <remarks> Docs: <see href="https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.readonlyrecommended"/> </remarks>
		[SupportByVersion("PowerPoint", 16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ReadOnlyRecommended
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ReadOnlyRecommended");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FullName"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Name"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Path"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Saved"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Saved
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Saved");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Saved", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.LayoutDirection"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpDirection LayoutDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpDirection>(this, "LayoutDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOptions"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PrintOptions PrintOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PrintOptions>(this, "PrintOptions", NetOffice.PowerPointApi.PrintOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Container"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Container
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.DisplayComments"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplayComments
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "DisplayComments");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FarEastLineBreakLevel"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel FarEastLineBreakLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel>(this, "FarEastLineBreakLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FarEastLineBreakLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.NoLineBreakBefore"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string NoLineBreakBefore
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NoLineBreakBefore");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoLineBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.NoLineBreakAfter"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string NoLineBreakAfter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NoLineBreakAfter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoLineBreakAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SlideShowWindow"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowWindow SlideShowWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowWindow>(this, "SlideShowWindow", NetOffice.PowerPointApi.SlideShowWindow.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FarEastLineBreakLanguage"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID>(this, "FarEastLineBreakLanguage");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FarEastLineBreakLanguage", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.DefaultLanguageID"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID DefaultLanguageID
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "DefaultLanguageID");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultLanguageID", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CommandBars"/> </remarks>
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
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PublishObjects PublishObjects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PublishObjects>(this, "PublishObjects", NetOffice.PowerPointApi.PublishObjects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.WebOptions WebOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.WebOptions>(this, "WebOptions", NetOffice.PowerPointApi.WebOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProject>(this, "HTMLProject", NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.EnvelopeVisible"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState EnvelopeVisible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "EnvelopeVisible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EnvelopeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.VBASigned"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState VBASigned
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "VBASigned");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SnapToGrid"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState SnapToGrid
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SnapToGrid");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.GridDistance"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single GridDistance
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridDistance");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridDistance", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Designs"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Designs Designs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Designs>(this, "Designs", NetOffice.PowerPointApi.Designs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Signatures"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSet>(this, "Signatures", NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.RemovePersonalInformation"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState RemovePersonalInformation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "RemovePersonalInformation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RemovePersonalInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpRevisionInfo HasRevisionInfo
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpRevisionInfo>(this, "HasRevisionInfo");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PasswordEncryptionProvider"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PasswordEncryptionProvider");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PasswordEncryptionAlgorithm"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PasswordEncryptionAlgorithm");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PasswordEncryptionKeyLength"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PasswordEncryptionKeyLength");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PasswordEncryptionFileProperties"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasswordEncryptionFileProperties");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Password"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string Password
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Password");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Password", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.WritePassword"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string WritePassword
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WritePassword");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WritePassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Permission"/> </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Permission>(this, "Permission", NetOffice.OfficeApi.Permission.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SharedWorkspace"/> </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Sync"/> </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Sync>(this, "Sync", NetOffice.OfficeApi.Sync.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.DocumentLibraryVersions"/> </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ContentTypeProperties"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MetaProperties>(this, "ContentTypeProperties", NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Int32 SectionCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SectionCount");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool HasSections
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasSections");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ServerPolicy"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ServerPolicy>(this, "ServerPolicy", NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.DocumentInspectors"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentInspectors>(this, "DocumentInspectors", NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.HasVBProject"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasVBProject");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CustomXMLParts"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLParts>(this, "CustomXMLParts", NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Final"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool Final
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Final");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Final", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CustomerData"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.CustomerData CustomerData
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.CustomerData>(this, "CustomerData", NetOffice.PowerPointApi.CustomerData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Research"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Research Research
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Research>(this, "Research", NetOffice.PowerPointApi.Research.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.EncryptionProvider"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public string EncryptionProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EncryptionProvider");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EncryptionProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SectionProperties"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.SectionProperties SectionProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SectionProperties>(this, "SectionProperties", NetOffice.PowerPointApi.SectionProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Coauthoring"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Coauthoring Coauthoring
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Coauthoring>(this, "Coauthoring", NetOffice.PowerPointApi.Coauthoring.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.InMergeMode"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool InMergeMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InMergeMode");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Broadcast"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Broadcast Broadcast
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Broadcast>(this, "Broadcast", NetOffice.PowerPointApi.Broadcast.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.HasNotesMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasNotesMaster
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasNotesMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.HasHandoutMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasHandoutMaster
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasHandoutMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideoStatus"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpMediaTaskStatus CreateVideoStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpMediaTaskStatus>(this, "CreateVideoStatus");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.chartdatapointtrack"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.guides"/> </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Guides Guides
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Guides>(this, "Guides", NetOffice.PowerPointApi.Guides.LateBindingApiWrapperType);
			}
		}
		
		/// <summary>
		/// True if the edits in the Presentation are automatically saved. Read/write Boolean.
		/// 
		/// SupportByVersion PowerPoint 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://learn.microsoft.com/fr-fr/office/vba/api/powerpoint.presentation.autosaveon"/> </remarks>
		[SupportByVersion("PowerPoint", 16)]
		public bool AutoSaveOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoSaveOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoSaveOn", value);
			}
		}



		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.AddTitleMaster"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master AddTitleMaster()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.PowerPointApi._Master>(this, "AddTitleMaster");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ApplyTemplate"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ApplyTemplate(string fileName)
		{
			 Factory.ExecuteMethod(this, "ApplyTemplate", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.NewWindow"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindow NewWindow()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DocumentWindow>(this, "NewWindow", NetOffice.PowerPointApi.DocumentWindow.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.FollowHyperlink"/> </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.AddToFavorites"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			 Factory.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Unused()
		{
			 Factory.ExecuteMethod(this, "Unused");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOut"/> </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		/// <param name="collate">optional NetOffice.OfficeApi.Enums.MsoTriState Collate = -99</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, printToFile, copies, collate });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOut"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOut"/> </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOut"/> </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOut"/> </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, printToFile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PrintOut"/> </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, printToFile, copies);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Save"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveAs"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName, object fileFormat, object embedTrueTypeFonts)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName, fileFormat, embedTrueTypeFonts);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveAs"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveAs"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveCopyAs"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName, object fileFormat, object embedTrueTypeFonts)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat, embedTrueTypeFonts);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveCopyAs"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveCopyAs"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Export"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		/// <param name="scaleHeight">optional Int32 ScaleHeight = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName, object scaleWidth, object scaleHeight)
		{
			 Factory.ExecuteMethod(this, "Export", path, filterName, scaleWidth, scaleHeight);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Export"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName)
		{
			 Factory.ExecuteMethod(this, "Export", path, filterName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Export"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName, object scaleWidth)
		{
			 Factory.ExecuteMethod(this, "Export", path, filterName, scaleWidth);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Close"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SetUndoText(string text)
		{
			 Factory.ExecuteMethod(this, "SetUndoText", text);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.UpdateLinks"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void UpdateLinks()
		{
			 Factory.ExecuteMethod(this, "UpdateLinks");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			 Factory.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cp">NetOffice.OfficeApi.Enums.MsoEncoding cp</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding cp)
		{
			 Factory.ExecuteMethod(this, "ReloadAs", cp);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="isDesignTemplate">NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void MakeIntoTemplate(NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate)
		{
			 Factory.ExecuteMethod(this, "MakeIntoTemplate", isDesignTemplate);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			 Factory.ExecuteMethod(this, "sblt", s);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.merge"/> </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void Merge(string path)
		{
			 Factory.ExecuteMethod(this, "Merge", path);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckIn"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckIn"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			 Factory.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckIn"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckIn"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CanCheckIn"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public bool CanCheckIn()
		{
			return Factory.ExecuteBoolMethodGet(this, "CanCheckIn");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		/// <param name="includeAttachment">optional object includeAttachment</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage, includeAttachment);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			 Factory.ExecuteMethod(this, "SendForReview");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			 Factory.ExecuteMethod(this, "ReplyWithChanges", showMessage);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			 Factory.ExecuteMethod(this, "ReplyWithChanges");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.EndReview"/> </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void EndReview()
		{
			 Factory.ExecuteMethod(this, "EndReview");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional string FileName = </param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void AddBaseline(object fileName)
		{
			 Factory.ExecuteMethod(this, "AddBaseline", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void AddBaseline()
		{
			 Factory.ExecuteMethod(this, "AddBaseline");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void RemoveBaseline()
		{
			 Factory.ExecuteMethod(this, "RemoveBaseline");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SetPasswordEncryptionOptions"/> </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">bool passwordEncryptionFileProperties</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, bool passwordEncryptionFileProperties)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SendFaxOverInternet"/> </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = false</param>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SendFaxOverInternet"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SendFaxOverInternet"/> </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SendFaxOverInternet"/> </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="afterSlide">bool afterSlide</param>
		/// <param name="sectionTitle">string sectionTitle</param>
		/// <param name="newSectionIndex">Int32 newSectionIndex</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void DeleteSection(Int32 index)
		{
			 Factory.ExecuteMethod(this, "DeleteSection", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void DisableSections()
		{
			 Factory.ExecuteMethod(this, "DisableSections");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public string sectionTitle(Int32 index)
		{
			return Factory.ExecuteStringMethodGet(this, "sectionTitle", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.RemoveDocumentInformation"/> </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type)
		{
			 Factory.ExecuteMethod(this, "RemoveDocumentInformation", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic, versionType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckInWithVersion"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CheckInWithVersion"/> </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		/// <param name="externalExporter">optional object externalExporter</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", path, fixedFormatType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", path, fixedFormatType, intent);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", path, fixedFormatType, intent, frameSlides);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ExportAsFixedFormat"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.GetWorkflowTasks"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTasks>(this, "GetWorkflowTasks", NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.GetWorkflowTemplates"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTemplates>(this, "GetWorkflowTemplates", NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.LockServerFile"/> </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void LockServerFile()
		{
			 Factory.ExecuteMethod(this, "LockServerFile");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.ApplyTheme"/> </remarks>
		/// <param name="themeName">string themeName</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ApplyTheme(string themeName)
		{
			 Factory.ExecuteMethod(this, "ApplyTheme", themeName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PublishSlides"/> </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		/// <param name="useSlideOrder">optional bool UseSlideOrder = false</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl, object overwrite, object useSlideOrder)
		{
			 Factory.ExecuteMethod(this, "PublishSlides", slideLibraryUrl, overwrite, useSlideOrder);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PublishSlides"/> </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl)
		{
			 Factory.ExecuteMethod(this, "PublishSlides", slideLibraryUrl);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.PublishSlides"/> </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl, object overwrite)
		{
			 Factory.ExecuteMethod(this, "PublishSlides", slideLibraryUrl, overwrite);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void Convert()
		{
			 Factory.ExecuteMethod(this, "Convert");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.MergeWithBaseline"/> </remarks>
		/// <param name="withPresentation">string withPresentation</param>
		/// <param name="baselinePresentation">string baselinePresentation</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void MergeWithBaseline(string withPresentation, string baselinePresentation)
		{
			 Factory.ExecuteMethod(this, "MergeWithBaseline", withPresentation, baselinePresentation);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.AcceptAll"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void AcceptAll()
		{
			 Factory.ExecuteMethod(this, "AcceptAll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.RejectAll"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void RejectAll()
		{
			 Factory.ExecuteMethod(this, "RejectAll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.EnsureAllMediaUpgraded"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void EnsureAllMediaUpgraded()
		{
			 Factory.ExecuteMethod(this, "EnsureAllMediaUpgraded");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.Convert2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Convert2(string fileName)
		{
			 Factory.ExecuteMethod(this, "Convert2", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideo"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		/// <param name="quality">optional Int32 Quality = 85</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond, object quality)
		{
			 Factory.ExecuteMethod(this, "CreateVideo", new object[]{ fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond, quality });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideo"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName)
		{
			 Factory.ExecuteMethod(this, "CreateVideo", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideo"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations)
		{
			 Factory.ExecuteMethod(this, "CreateVideo", fileName, useTimingsAndNarrations);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideo"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration)
		{
			 Factory.ExecuteMethod(this, "CreateVideo", fileName, useTimingsAndNarrations, defaultSlideDuration);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideo"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution)
		{
			 Factory.ExecuteMethod(this, "CreateVideo", fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.CreateVideo"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond)
		{
			 Factory.ExecuteMethod(this, "CreateVideo", new object[]{ fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.applytemplate2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="variant">string variant</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ApplyTemplate2(string fileName, string variant)
		{
			 Factory.ExecuteMethod(this, "ApplyTemplate2", fileName, variant);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		/// <param name="externalExporter">optional object externalExporter</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object includeMarkup, object externalExporter)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, includeMarkup, externalExporter });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", path, fixedFormatType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", path, fixedFormatType, intent);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", path, fixedFormatType, intent, frameSlides);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.presentation.exportasfixedformat2"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
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
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object includeMarkup)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, includeMarkup });
		}

		#endregion

		#pragma warning restore
	}
}
