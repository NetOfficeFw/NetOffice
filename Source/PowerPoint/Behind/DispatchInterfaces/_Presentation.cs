using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface _Presentation 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Presentation : COMObject, NetOffice.PowerPointApi._Presentation
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
                    _contractType = typeof(NetOffice.PowerPointApi._Presentation);
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
                    _type = typeof(_Presentation);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Presentation() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745080.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743905.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745484.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master SlideMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "SlideMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746378.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master TitleMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "TitleMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745657.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasTitleMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasTitleMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744870.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string TemplateName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TemplateName");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743938.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master NotesMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "NotesMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746405.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master HandoutMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.PowerPointApi._Master>(this, "HandoutMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746142.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Slides Slides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Slides>(this, "Slides", typeof(NetOffice.PowerPointApi.Slides));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745413.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PageSetup PageSetup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PageSetup>(this, "PageSetup", typeof(NetOffice.PowerPointApi.PageSetup));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744763.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ColorSchemes ColorSchemes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ColorSchemes>(this, "ColorSchemes", typeof(NetOffice.PowerPointApi.ColorSchemes));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746216.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ExtraColors ExtraColors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ExtraColors>(this, "ExtraColors", typeof(NetOffice.PowerPointApi.ExtraColors));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745621.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowSettings SlideShowSettings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowSettings>(this, "SlideShowSettings", typeof(NetOffice.PowerPointApi.SlideShowSettings));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744620.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Fonts Fonts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Fonts>(this, "Fonts", typeof(NetOffice.PowerPointApi.Fonts));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746292.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindows Windows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DocumentWindows>(this, "Windows", typeof(NetOffice.PowerPointApi.DocumentWindows));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744602.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Tags Tags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Tags>(this, "Tags", typeof(NetOffice.PowerPointApi.Tags));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744397.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape DefaultShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "DefaultShape", typeof(NetOffice.PowerPointApi.Shape));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746376.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object BuiltInDocumentProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BuiltInDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744661.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object CustomDocumentProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745299.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", typeof(NetOffice.VBIDEApi.VBProject));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746329.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState ReadOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746313.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745890.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745125.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744884.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Saved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Saved");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Saved", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744109.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpDirection LayoutDirection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpDirection>(this, "LayoutDirection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LayoutDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744540.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PrintOptions PrintOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PrintOptions>(this, "PrintOptions", typeof(NetOffice.PowerPointApi.PrintOptions));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746277.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Container
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745979.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplayComments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "DisplayComments");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746803.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel FarEastLineBreakLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel>(this, "FarEastLineBreakLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FarEastLineBreakLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746404.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string NoLineBreakBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NoLineBreakBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoLineBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746110.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string NoLineBreakAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NoLineBreakAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoLineBreakAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745765.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowWindow SlideShowWindow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowWindow>(this, "SlideShowWindow", typeof(NetOffice.PowerPointApi.SlideShowWindow));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746489.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID>(this, "FarEastLineBreakLanguage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FarEastLineBreakLanguage", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745465.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID DefaultLanguageID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "DefaultLanguageID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultLanguageID", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746786.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PublishObjects>(this, "PublishObjects", typeof(NetOffice.PowerPointApi.PublishObjects));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.WebOptions>(this, "WebOptions", typeof(NetOffice.PowerPointApi.WebOptions));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProject>(this, "HTMLProject", typeof(NetOffice.OfficeApi.HTMLProject));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746516.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState EnvelopeVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "EnvelopeVisible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EnvelopeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746671.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState VBASigned
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "VBASigned");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746323.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState SnapToGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "SnapToGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744975.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Single GridDistance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridDistance");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridDistance", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744959.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Designs Designs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Designs>(this, "Designs", typeof(NetOffice.PowerPointApi.Designs));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745705.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSet>(this, "Signatures", typeof(NetOffice.OfficeApi.SignatureSet));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746134.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState RemovePersonalInformation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "RemovePersonalInformation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RemovePersonalInformation", value);
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
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpRevisionInfo>(this, "HasRevisionInfo");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743904.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PasswordEncryptionProvider");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745251.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PasswordEncryptionAlgorithm");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744792.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PasswordEncryptionKeyLength");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743937.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PasswordEncryptionFileProperties");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745703.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string Password
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Password");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Password", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744704.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public string WritePassword
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WritePassword");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WritePassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744658.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Permission>(this, "Permission", typeof(NetOffice.OfficeApi.Permission));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745343.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", typeof(NetOffice.OfficeApi.SharedWorkspace));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745948.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Sync>(this, "Sync", typeof(NetOffice.OfficeApi.Sync));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744818.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", typeof(NetOffice.OfficeApi.DocumentLibraryVersions));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745118.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MetaProperties>(this, "ContentTypeProperties", typeof(NetOffice.OfficeApi.MetaProperties));
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
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SectionCount");
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
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasSections");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745108.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ServerPolicy>(this, "ServerPolicy", typeof(NetOffice.OfficeApi.ServerPolicy));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744654.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentInspectors>(this, "DocumentInspectors", typeof(NetOffice.OfficeApi.DocumentInspectors));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746792.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasVBProject");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745253.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLParts>(this, "CustomXMLParts", typeof(NetOffice.OfficeApi.CustomXMLParts));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743879.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public bool Final
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Final");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Final", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745858.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.CustomerData CustomerData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.CustomerData>(this, "CustomerData", typeof(NetOffice.PowerPointApi.CustomerData));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746518.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Research Research
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Research>(this, "Research", typeof(NetOffice.PowerPointApi.Research));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745747.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public string EncryptionProvider
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EncryptionProvider");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EncryptionProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744806.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.SectionProperties SectionProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SectionProperties>(this, "SectionProperties", typeof(NetOffice.PowerPointApi.SectionProperties));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745326.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Coauthoring Coauthoring
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Coauthoring>(this, "Coauthoring", typeof(NetOffice.PowerPointApi.Coauthoring));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746423.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool InMergeMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InMergeMode");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744901.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Broadcast Broadcast
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Broadcast>(this, "Broadcast", typeof(NetOffice.PowerPointApi.Broadcast));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745775.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasNotesMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasNotesMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744680.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasHandoutMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasHandoutMaster");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743993.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpMediaTaskStatus CreateVideoStatus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpMediaTaskStatus>(this, "CreateVideoStatus");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229098.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public bool ChartDataPointTrack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ChartDataPointTrack");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ChartDataPointTrack", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229702.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Guides Guides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Guides>(this, "Guides", typeof(NetOffice.PowerPointApi.Guides));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746001.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.PowerPointApi._Master AddTitleMaster()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.PowerPointApi._Master>(this, "AddTitleMaster");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743876.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ApplyTemplate(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyTemplate", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744342.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.DocumentWindow NewWindow()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.DocumentWindow>(this, "NewWindow", typeof(NetOffice.PowerPointApi.DocumentWindow));
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744969.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Unused()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Unused");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		/// <param name="collate">optional NetOffice.OfficeApi.Enums.MsoTriState Collate = -99</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies, object collate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[]{ from, to, printToFile, copies, collate });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to, printToFile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object printToFile, object copies)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to, printToFile, copies);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745194.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName, object fileFormat, object embedTrueTypeFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName, fileFormat, embedTrueTypeFonts);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveAs(string fileName, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName, object fileFormat, object embedTrueTypeFonts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat, embedTrueTypeFonts);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(string fileName, object fileFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", fileName, fileFormat);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		/// <param name="scaleHeight">optional Int32 ScaleHeight = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName, object scaleWidth, object scaleHeight)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", path, filterName, scaleWidth, scaleHeight);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", path, filterName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Export(string path, string filterName, object scaleWidth)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", path, filterName, scaleWidth);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743857.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SetUndoText(string text)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetUndoText", text);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744167.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void UpdateLinks()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateLinks");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cp">NetOffice.OfficeApi.Enums.MsoEncoding cp</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding cp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReloadAs", cp);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="isDesignTemplate">NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void MakeIntoTemplate(NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MakeIntoTemplate", isDesignTemplate);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "sblt", s);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228409.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void Merge(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Merge", path);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744274.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public bool CanCheckIn()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckIn");
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage, includeAttachment);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyWithChanges", showMessage);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyWithChanges");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746226.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void EndReview()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndReview");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional string FileName = </param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void AddBaseline(object fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddBaseline", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void AddBaseline()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddBaseline");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void RemoveBaseline()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveBaseline");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743880.aspx </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">bool passwordEncryptionFileProperties</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, bool passwordEncryptionFileProperties)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = false</param>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject);
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteSection", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void DisableSections()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DisableSections");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public string sectionTitle(Int32 index)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "sectionTitle", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744345.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveDocumentInformation", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic, versionType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, externalExporter });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", path, fixedFormatType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", path, fixedFormatType, intent);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", path, fixedFormatType, intent, frameSlides);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746373.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTasks>(this, "GetWorkflowTasks", typeof(NetOffice.OfficeApi.WorkflowTasks));
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746712.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTemplates>(this, "GetWorkflowTemplates", typeof(NetOffice.OfficeApi.WorkflowTemplates));
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744831.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void LockServerFile()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LockServerFile");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746528.aspx </remarks>
		/// <param name="themeName">string themeName</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void ApplyTheme(string themeName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyTheme", themeName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		/// <param name="useSlideOrder">optional bool UseSlideOrder = false</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl, object overwrite, object useSlideOrder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PublishSlides", slideLibraryUrl, overwrite, useSlideOrder);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PublishSlides", slideLibraryUrl);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void PublishSlides(string slideLibraryUrl, object overwrite)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PublishSlides", slideLibraryUrl, overwrite);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void Convert()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Convert");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744063.aspx </remarks>
		/// <param name="withPresentation">string withPresentation</param>
		/// <param name="baselinePresentation">string baselinePresentation</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void MergeWithBaseline(string withPresentation, string baselinePresentation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MergeWithBaseline", withPresentation, baselinePresentation);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745418.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void AcceptAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745993.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void RejectAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744528.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void EnsureAllMediaUpgraded()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EnsureAllMediaUpgraded");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743830.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void Convert2(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Convert2", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		/// <param name="quality">optional Int32 Quality = 85</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond, object quality)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateVideo", new object[]{ fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond, quality });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateVideo", fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateVideo", fileName, useTimingsAndNarrations);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateVideo", fileName, useTimingsAndNarrations, defaultSlideDuration);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateVideo", fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateVideo", new object[]{ fileName, useTimingsAndNarrations, defaultSlideDuration, vertResolution, framesPerSecond });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228125.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="variant">string variant</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ApplyTemplate2(string fileName, string variant)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyTemplate2", fileName, variant);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, includeMarkup, externalExporter });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", path, fixedFormatType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", path, fixedFormatType, intent);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", path, fixedFormatType, intent, frameSlides);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
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
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat2", new object[]{ path, fixedFormatType, intent, frameSlides, handoutOrder, outputType, printHiddenSlides, printRange, rangeType, slideShowName, includeDocProperties, keepIRMSettings, docStructureTags, bitmapMissingFonts, useISO19005_1, includeMarkup });
		}

		#endregion

		#pragma warning restore
	}
}


