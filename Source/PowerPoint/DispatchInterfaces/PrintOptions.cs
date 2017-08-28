using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PrintOptions 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744130.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PrintOptions : COMObject
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
                    _type = typeof(PrintOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PrintOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PrintOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PrintOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PrintOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PrintOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PrintOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PrintOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PrintOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745165.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746721.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpPrintColorType PrintColorType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpPrintColorType>(this, "PrintColorType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintColorType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744827.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Collate
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Collate");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Collate", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744894.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState FitToPage
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FitToPage");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FitToPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744843.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState FrameSlides
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FrameSlides");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FrameSlides", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745110.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 NumberOfCopies
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NumberOfCopies");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberOfCopies", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745124.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpPrintOutputType>(this, "OutputType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OutputType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746418.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744580.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "PrintHiddenSlides");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintHiddenSlides", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746359.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState PrintInBackground
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "PrintInBackground");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintInBackground", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744876.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpPrintRangeType>(this, "RangeType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RangeType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746322.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.PrintRanges Ranges
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PrintRanges>(this, "Ranges", NetOffice.PowerPointApi.PrintRanges.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746747.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState PrintFontsAsGraphics
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "PrintFontsAsGraphics");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintFontsAsGraphics", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744011.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public string SlideShowName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SlideShowName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SlideShowName", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744710.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746393.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder>(this, "HandoutOrder");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HandoutOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744618.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState PrintComments
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "PrintComments");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746364.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 sectionIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "sectionIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "sectionIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745002.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HighQuality
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HighQuality");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HighQuality", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
