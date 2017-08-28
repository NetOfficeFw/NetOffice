using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Options 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822397.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Options : COMObject
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
                    _type = typeof(Options);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Options(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Options(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Options(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193097.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839140.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822677.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197714.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowAccentedUppercase
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowAccentedUppercase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowAccentedUppercase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool WPHelp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WPHelp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WPHelp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool WPDocNavKeys
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WPDocNavKeys");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WPDocNavKeys", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822573.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Pagination
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Pagination");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Pagination", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool BlueScreen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BlueScreen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BlueScreen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838088.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool EnableSound
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableSound");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableSound", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195314.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ConfirmConversions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConfirmConversions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConfirmConversions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845219.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UpdateLinksAtOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateLinksAtOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateLinksAtOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839735.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SendMailAttach
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SendMailAttach");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SendMailAttach", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821587.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMeasurementUnits MeasurementUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMeasurementUnits>(this, "MeasurementUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MeasurementUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197267.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ButtonFieldClicks
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ButtonFieldClicks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ButtonFieldClicks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822969.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShortMenuNames
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShortMenuNames");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortMenuNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231613.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool RTFInClipboard
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RTFInClipboard");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RTFInClipboard", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845103.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UpdateFieldsAtPrint
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateFieldsAtPrint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateFieldsAtPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195196.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintProperties");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintProperties", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840788.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintFieldCodes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintFieldCodes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintFieldCodes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822925.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintComments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintComments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195502.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintHiddenText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintHiddenText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintHiddenText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834946.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool EnvelopeFeederInstalled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnvelopeFeederInstalled");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194722.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UpdateLinksAtPrint
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateLinksAtPrint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateLinksAtPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194191.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintBackground
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintBackground");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintBackground", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193448.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintDrawingObjects
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintDrawingObjects");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintDrawingObjects", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839854.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string DefaultTray
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultTray");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultTray", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193880.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 DefaultTrayID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DefaultTrayID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultTrayID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844899.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CreateBackup
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CreateBackup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CreateBackup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowFastSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFastSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFastSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839103.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SavePropertiesPrompt
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SavePropertiesPrompt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SavePropertiesPrompt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837470.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SaveNormalPrompt
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveNormalPrompt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveNormalPrompt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845537.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SaveInterval
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SaveInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835780.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool BackgroundSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackgroundSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197884.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdInsertedTextMark InsertedTextMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdInsertedTextMark>(this, "InsertedTextMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InsertedTextMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838549.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDeletedTextMark DeletedTextMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDeletedTextMark>(this, "DeletedTextMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DeletedTextMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840041.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisedLinesMark RevisedLinesMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisedLinesMark>(this, "RevisedLinesMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisedLinesMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195698.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex InsertedTextColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "InsertedTextColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InsertedTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839343.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex DeletedTextColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "DeletedTextColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DeletedTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837475.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex RevisedLinesColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "RevisedLinesColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisedLinesColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193744.aspx </remarks>
		/// <param name="path">NetOffice.WordApi.Enums.WdDefaultFilePath path</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_DefaultFilePath(NetOffice.WordApi.Enums.WdDefaultFilePath path)
		{
			return Factory.ExecuteStringPropertyGet(this, "DefaultFilePath", path);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="path">NetOffice.WordApi.Enums.WdDefaultFilePath path</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_DefaultFilePath(NetOffice.WordApi.Enums.WdDefaultFilePath path, string value)
		{
			Factory.ExecutePropertySet(this, "DefaultFilePath", path, value);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DefaultFilePath
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193744.aspx </remarks>
		/// <param name="path">NetOffice.WordApi.Enums.WdDefaultFilePath path</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_DefaultFilePath")]
		public string DefaultFilePath(NetOffice.WordApi.Enums.WdDefaultFilePath path)
		{
			return get_DefaultFilePath(path);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192407.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Overtype
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Overtype");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Overtype", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838563.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ReplaceSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReplaceSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReplaceSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840488.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowDragAndDrop
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDragAndDrop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDragAndDrop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoWordSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoWordSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoWordSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835473.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool INSKeyForPaste
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "INSKeyForPaste");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "INSKeyForPaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196264.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SmartCutPaste
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SmartCutPaste");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmartCutPaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192127.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool TabIndentKey
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TabIndentKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabIndentKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837511.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string PictureEditor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PictureEditor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureEditor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823217.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AnimateScreenMovements
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AnimateScreenMovements");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AnimateScreenMovements", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool VirusProtection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VirusProtection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VirusProtection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836037.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisedPropertiesMark RevisedPropertiesMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisedPropertiesMark>(this, "RevisedPropertiesMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisedPropertiesMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844801.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex RevisedPropertiesColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "RevisedPropertiesColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisedPropertiesColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192408.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SnapToGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SnapToGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapToGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820872.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SnapToShapes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SnapToShapes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SnapToShapes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191977.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridDistanceHorizontal
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridDistanceHorizontal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridDistanceHorizontal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198031.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridDistanceVertical
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridDistanceVertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridDistanceVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836862.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridOriginHorizontal
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridOriginHorizontal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridOriginHorizontal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197259.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single GridOriginVertical
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "GridOriginVertical");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridOriginVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840173.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool InlineConversion
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InlineConversion");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InlineConversion", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845818.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IMEAutomaticControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IMEAutomaticControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IMEAutomaticControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834928.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatApplyHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatApplyHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatApplyHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840580.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatApplyLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatApplyLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatApplyLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835804.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatApplyBulletedLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatApplyBulletedLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatApplyBulletedLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837023.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatApplyOtherParas
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatApplyOtherParas");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatApplyOtherParas", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192381.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplaceQuotes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplaceQuotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplaceQuotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196304.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplaceSymbols
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplaceSymbols");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplaceSymbols", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821610.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplaceOrdinals
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplaceOrdinals");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplaceOrdinals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839714.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplaceFractions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplaceFractions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplaceFractions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835414.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplacePlainTextEmphasis
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplacePlainTextEmphasis");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplacePlainTextEmphasis", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838290.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatPreserveStyles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatPreserveStyles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatPreserveStyles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191857.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197436.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyBorders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyBorders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyBorders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196813.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyBulletedLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyBulletedLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyBulletedLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835487.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyNumberedLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyNumberedLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyNumberedLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838540.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceQuotes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceQuotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceQuotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845135.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceSymbols
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceSymbols");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceSymbols", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840187.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceOrdinals
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceOrdinals");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceOrdinals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840985.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceFractions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceFractions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceFractions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821524.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplacePlainTextEmphasis
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplacePlainTextEmphasis");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplacePlainTextEmphasis", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821823.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeFormatListItemBeginning
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeFormatListItemBeginning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeFormatListItemBeginning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191698.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeDefineStyles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeDefineStyles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeDefineStyles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822538.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatPlainTextWordMail
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatPlainTextWordMail");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatPlainTextWordMail", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834579.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceHyperlinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceHyperlinks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceHyperlinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836574.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplaceHyperlinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplaceHyperlinks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplaceHyperlinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845670.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex DefaultHighlightColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "DefaultHighlightColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultHighlightColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197541.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLineStyle DefaultBorderLineStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLineStyle>(this, "DefaultBorderLineStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultBorderLineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822988.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckSpellingAsYouType
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckSpellingAsYouType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckSpellingAsYouType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845689.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckGrammarAsYouType
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckGrammarAsYouType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckGrammarAsYouType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193102.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IgnoreInternetAndFileAddresses
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreInternetAndFileAddresses");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreInternetAndFileAddresses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193126.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowReadabilityStatistics
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowReadabilityStatistics");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowReadabilityStatistics", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195498.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IgnoreUppercase
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreUppercase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreUppercase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193439.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IgnoreMixedDigits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreMixedDigits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreMixedDigits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839090.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SuggestFromMainDictionaryOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SuggestFromMainDictionaryOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SuggestFromMainDictionaryOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192763.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SuggestSpellingCorrections
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SuggestSpellingCorrections");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SuggestSpellingCorrections", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836105.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLineWidth DefaultBorderLineWidth
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLineWidth>(this, "DefaultBorderLineWidth");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultBorderLineWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836612.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckGrammarWithSpelling
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckGrammarWithSpelling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckGrammarWithSpelling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822937.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOpenFormat DefaultOpenFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOpenFormat>(this, "DefaultOpenFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultOpenFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192369.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintDraft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintDraft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintDraft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837504.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintReverse
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintReverse");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintReverse", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836101.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MapPaperSize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MapPaperSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MapPaperSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191733.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyTables
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyTables");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyTables", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837932.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatApplyFirstIndents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatApplyFirstIndents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatApplyFirstIndents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840121.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatMatchParentheses
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatMatchParentheses");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatMatchParentheses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193385.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatReplaceFarEastDashes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatReplaceFarEastDashes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatReplaceFarEastDashes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194760.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatDeleteAutoSpaces
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatDeleteAutoSpaces");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatDeleteAutoSpaces", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838899.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyFirstIndents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyFirstIndents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyFirstIndents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836758.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyDates
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyDates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyDates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191743.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyClosings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyClosings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyClosings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840598.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeMatchParentheses
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeMatchParentheses");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeMatchParentheses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192183.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceFarEastDashes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceFarEastDashes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceFarEastDashes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835424.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeDeleteAutoSpaces
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeDeleteAutoSpaces");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeDeleteAutoSpaces", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822989.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeInsertClosings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeInsertClosings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeInsertClosings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837521.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeAutoLetterWizard
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeAutoLetterWizard");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeAutoLetterWizard", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839749.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeInsertOvers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeInsertOvers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeInsertOvers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836939.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisplayGridLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayGridLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayGridLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193080.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyCase
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyCase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyCase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834822.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyByte
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyByte");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyByte", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821136.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyHiragana
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyHiragana");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyHiragana", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820876.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzySmallKana
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzySmallKana");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzySmallKana", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845789.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyDash
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyDash");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyDash", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192390.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyIterationMark
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyIterationMark");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyIterationMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197979.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyKanji
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyKanji");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyKanji", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197574.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyOldKana
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyOldKana");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyOldKana", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840023.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyProlongedSoundMark
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyProlongedSoundMark");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyProlongedSoundMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194736.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyDZ
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyDZ");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyDZ", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193404.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyBV
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyBV");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyBV", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835167.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyTC
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyTC");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyTC", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840893.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyHF
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyHF");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyHF", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823220.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyZJ
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyZJ");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyZJ", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840784.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyAY
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyAY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyAY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192829.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyKiKu
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyKiKu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyKiKu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839895.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzyPunctuation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzyPunctuation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzyPunctuation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836852.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzySpace
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzySpace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzySpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836589.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ApplyFarEastFontsToAscii
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyFarEastFontsToAscii");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyFarEastFontsToAscii", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839001.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ConvertHighAnsiToFarEast
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConvertHighAnsiToFarEast");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConvertHighAnsiToFarEast", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837907.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintOddPagesInAscendingOrder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintOddPagesInAscendingOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintOddPagesInAscendingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193422.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool PrintEvenPagesInAscendingOrder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintEvenPagesInAscendingOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintEvenPagesInAscendingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822955.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex DefaultBorderColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "DefaultBorderColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultBorderColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195918.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool EnableMisusedWordsDictionary
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableMisusedWordsDictionary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableMisusedWordsDictionary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837956.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowCombinedAuxiliaryForms
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowCombinedAuxiliaryForms");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowCombinedAuxiliaryForms", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193700.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HangulHanjaFastConversion
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HangulHanjaFastConversion");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HangulHanjaFastConversion", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840939.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CheckHangulEndings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckHangulEndings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckHangulEndings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192761.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool EnableHangulHanjaRecentOrdering
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableHangulHanjaRecentOrdering");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableHangulHanjaRecentOrdering", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194480.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMultipleWordConversionsMode MultipleWordConversionsMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMultipleWordConversionsMode>(this, "MultipleWordConversionsMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MultipleWordConversionsMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193704.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColor DefaultBorderColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "DefaultBorderColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultBorderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845681.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowPixelUnits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPixelUnits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPixelUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838350.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseCharacterUnit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseCharacterUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseCharacterUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821257.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowCompoundNounProcessing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowCompoundNounProcessing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowCompoundNounProcessing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192341.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AutoKeyboardSwitching
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoKeyboardSwitching");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoKeyboardSwitching", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196904.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDocumentViewDirection DocumentViewDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDocumentViewDirection>(this, "DocumentViewDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DocumentViewDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838701.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdArabicNumeral ArabicNumeral
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdArabicNumeral>(this, "ArabicNumeral");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ArabicNumeral", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192546.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMonthNames MonthNames
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMonthNames>(this, "MonthNames");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MonthNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840654.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCursorMovement CursorMovement
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCursorMovement>(this, "CursorMovement");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CursorMovement", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838721.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdVisualSelection VisualSelection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdVisualSelection>(this, "VisualSelection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VisualSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836603.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowDiacritics
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDiacritics");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDiacritics", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835406.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowControlCharacters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowControlCharacters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowControlCharacters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194559.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AddControlCharacters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AddControlCharacters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddControlCharacters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834902.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AddBiDirectionalMarksWhenSavingTextFile
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AddBiDirectionalMarksWhenSavingTextFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddBiDirectionalMarksWhenSavingTextFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191934.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool StrictInitialAlefHamza
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StrictInitialAlefHamza");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StrictInitialAlefHamza", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197250.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool StrictFinalYaa
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StrictFinalYaa");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StrictFinalYaa", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822686.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdHebSpellStart HebrewMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdHebSpellStart>(this, "HebrewMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HebrewMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840696.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdAraSpeller ArabicMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdAraSpeller>(this, "ArabicMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ArabicMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194350.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowClickAndTypeMouse
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowClickAndTypeMouse");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowClickAndTypeMouse", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196521.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseGermanSpellingReform
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseGermanSpellingReform");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseGermanSpellingReform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837698.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdHighAnsiText InterpretHighAnsi
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdHighAnsiText>(this, "InterpretHighAnsi");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InterpretHighAnsi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838139.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AddHebDoubleQuote
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AddHebDoubleQuote");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddHebDoubleQuote", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840949.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseDiffDiacColor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseDiffDiacColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseDiffDiacColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837461.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColor DiacriticColorVal
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "DiacriticColorVal");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DiacriticColorVal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197926.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool OptimizeForWord97byDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OptimizeForWord97byDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OptimizeForWord97byDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191787.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool LocalNetworkFile
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LocalNetworkFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LocalNetworkFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834589.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool TypeNReplace
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TypeNReplace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TypeNReplace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837742.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool SequenceCheck
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SequenceCheck");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SequenceCheck", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool BackgroundOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackgroundOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196309.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisableFeaturesbyDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableFeaturesbyDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableFeaturesbyDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192604.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteAdjustWordSpacing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteAdjustWordSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteAdjustWordSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845313.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteAdjustParagraphSpacing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteAdjustParagraphSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteAdjustParagraphSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822919.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteAdjustTableFormatting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteAdjustTableFormatting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteAdjustTableFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191988.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteSmartStyleBehavior
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteSmartStyleBehavior");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteSmartStyleBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196815.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteMergeFromPPT
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteMergeFromPPT");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteMergeFromPPT", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838529.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteMergeFromXL
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteMergeFromXL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteMergeFromXL", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192192.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool CtrlClickHyperlinkToOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CtrlClickHyperlinkToOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CtrlClickHyperlinkToOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837338.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdWrapTypeMerged PictureWrapType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdWrapTypeMerged>(this, "PictureWrapType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureWrapType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835978.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfterbyDefault
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdDisableFeaturesIntroducedAfter>(this, "DisableFeaturesIntroducedAfterbyDefault");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisableFeaturesIntroducedAfterbyDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838684.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteSmartCutPaste
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteSmartCutPaste");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteSmartCutPaste", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195680.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisplayPasteOptions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPasteOptions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPasteOptions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845093.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PromptUpdateStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PromptUpdateStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PromptUpdateStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191974.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public string DefaultEPostageApp
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultEPostageApp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultEPostageApp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845118.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoEncoding DefaultTextEncoding
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoEncoding>(this, "DefaultTextEncoding");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultTextEncoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool LabelSmartTags
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LabelSmartTags");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelSmartTags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisplaySmartTagButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplaySmartTagButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplaySmartTagButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194056.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool WarnBeforeSavingPrintingSendingMarkup
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WarnBeforeSavingPrintingSendingMarkup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WarnBeforeSavingPrintingSendingMarkup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198162.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool StoreRSIDOnSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StoreRSIDOnSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StoreRSIDOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197116.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowFormatError
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFormatError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFormatError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821002.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool FormatScanning
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FormatScanning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormatScanning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821974.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool PasteMergeLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteMergeLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteMergeLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838923.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoCreateNewDrawings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoCreateNewDrawings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoCreateNewDrawings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194031.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool SmartParaSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SmartParaSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmartParaSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836119.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisionsBalloonPrintOrientation RevisionsBalloonPrintOrientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsBalloonPrintOrientation>(this, "RevisionsBalloonPrintOrientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisionsBalloonPrintOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196601.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex CommentsColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "CommentsColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CommentsColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840314.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool PrintXMLTag
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintXMLTag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintXMLTag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821941.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool PrintBackgrounds
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintBackgrounds");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintBackgrounds", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837935.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool AllowReadingMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowReadingMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowReadingMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821638.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ShowMarkupOpenSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowMarkupOpenSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowMarkupOpenSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192409.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool SmartCursoring
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SmartCursoring");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmartCursoring", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193743.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMoveToTextMark MoveToTextMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMoveToTextMark>(this, "MoveToTextMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MoveToTextMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838132.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdMoveFromTextMark MoveFromTextMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMoveFromTextMark>(this, "MoveFromTextMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MoveFromTextMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839385.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string BibliographyStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BibliographyStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BibliographyStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192420.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string BibliographySort
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BibliographySort");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BibliographySort", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196386.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCellColor InsertedCellColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCellColor>(this, "InsertedCellColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InsertedCellColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197815.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCellColor DeletedCellColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCellColor>(this, "DeletedCellColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DeletedCellColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195592.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCellColor MergedCellColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCellColor>(this, "MergedCellColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MergedCellColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837659.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCellColor SplitCellColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCellColor>(this, "SplitCellColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitCellColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192432.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowSelectionFloaties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSelectionFloaties");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSelectionFloaties", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838284.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowMenuFloaties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowMenuFloaties");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowMenuFloaties", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194003.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowDevTools
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDevTools");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDevTools", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197724.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool EnableLivePreview
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableLivePreview");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableLivePreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193695.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool OMathAutoBuildUp
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OMathAutoBuildUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathAutoBuildUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool AlwaysUseClearType
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlwaysUseClearType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlwaysUseClearType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196208.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPasteOptions PasteFormatWithinDocument
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPasteOptions>(this, "PasteFormatWithinDocument");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PasteFormatWithinDocument", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195901.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPasteOptions PasteFormatBetweenDocuments
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPasteOptions>(this, "PasteFormatBetweenDocuments");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PasteFormatBetweenDocuments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191986.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPasteOptions PasteFormatBetweenStyledDocuments
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPasteOptions>(this, "PasteFormatBetweenStyledDocuments");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PasteFormatBetweenStyledDocuments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821957.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPasteOptions PasteFormatFromExternalSource
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPasteOptions>(this, "PasteFormatFromExternalSource");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PasteFormatFromExternalSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837864.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool PasteOptionKeepBulletsAndNumbers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasteOptionKeepBulletsAndNumbers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PasteOptionKeepBulletsAndNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194398.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool INSKeyForOvertype
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "INSKeyForOvertype");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "INSKeyForOvertype", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839700.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool RepeatWord
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RepeatWord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RepeatWord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845254.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFrenchSpeller FrenchReform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFrenchSpeller>(this, "FrenchReform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FrenchReform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838921.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ContextualSpeller
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ContextualSpeller");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ContextualSpeller", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845155.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex MoveToTextColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "MoveToTextColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MoveToTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845089.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex MoveFromTextColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "MoveFromTextColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MoveFromTextColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193754.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool OMathCopyLF
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OMathCopyLF");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OMathCopyLF", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198326.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool UseNormalStyleForList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseNormalStyleForList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseNormalStyleForList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837509.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool AllowOpenInDraftView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowOpenInDraftView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowOpenInDraftView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196200.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool EnableLegacyIMEMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableLegacyIMEMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableLegacyIMEMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834586.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool DoNotPromptForConvert
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DoNotPromptForConvert");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DoNotPromptForConvert", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837227.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool PrecisePositioning
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrecisePositioning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrecisePositioning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834581.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdUpdateStyleListBehavior UpdateStyleListBehavior
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdUpdateStyleListBehavior>(this, "UpdateStyleListBehavior");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "UpdateStyleListBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194566.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool StrictTaaMarboota
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StrictTaaMarboota");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StrictTaaMarboota", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836004.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool StrictRussianE
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StrictRussianE");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StrictRussianE", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822308.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdSpanishSpeller SpanishMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSpanishSpeller>(this, "SpanishMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SpanishMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195401.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdPortugueseReform PortugalReform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPortugueseReform>(this, "PortugalReform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PortugalReform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836096.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdPortugueseReform BrazilReform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPortugueseReform>(this, "BrazilReform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BrazilReform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822173.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool UpdateFieldsWithTrackedChangesAtPrint
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateFieldsWithTrackedChangesAtPrint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateFieldsWithTrackedChangesAtPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227416.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool DisplayAlignmentGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAlignmentGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAlignmentGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227290.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool PageAlignmentGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PageAlignmentGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageAlignmentGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227622.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool MarginAlignmentGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MarginAlignmentGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginAlignmentGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228294.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool ParagraphAlignmentGuides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ParagraphAlignmentGuides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ParagraphAlignmentGuides", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230717.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool EnableLiveDrag
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableLiveDrag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableLiveDrag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232153.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool UseSubPixelPositioning
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseSubPixelPositioning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseSubPixelPositioning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231044.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool AlertIfNotDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlertIfNotDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlertIfNotDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232151.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool EnableProofingToolsAdvertisement
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableProofingToolsAdvertisement");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableProofingToolsAdvertisement", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228644.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool PreferCloudSaveLocations
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PreferCloudSaveLocations");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PreferCloudSaveLocations", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232362.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool SkyDriveSignInOption
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SkyDriveSignInOption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SkyDriveSignInOption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231806.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool ExpandHeadingsOnOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ExpandHeadingsOnOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ExpandHeadingsOnOpen", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandKeyHelp">optional object commandKeyHelp</param>
		/// <param name="docNavigationKeys">optional object docNavigationKeys</param>
		/// <param name="mouseSimulation">optional object mouseSimulation</param>
		/// <param name="demoGuidance">optional object demoGuidance</param>
		/// <param name="demoSpeed">optional object demoSpeed</param>
		/// <param name="helpType">optional object helpType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions(object commandKeyHelp, object docNavigationKeys, object mouseSimulation, object demoGuidance, object demoSpeed, object helpType)
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions", new object[]{ commandKeyHelp, docNavigationKeys, mouseSimulation, demoGuidance, demoSpeed, helpType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions()
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandKeyHelp">optional object commandKeyHelp</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions(object commandKeyHelp)
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions", commandKeyHelp);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandKeyHelp">optional object commandKeyHelp</param>
		/// <param name="docNavigationKeys">optional object docNavigationKeys</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions(object commandKeyHelp, object docNavigationKeys)
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions", commandKeyHelp, docNavigationKeys);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandKeyHelp">optional object commandKeyHelp</param>
		/// <param name="docNavigationKeys">optional object docNavigationKeys</param>
		/// <param name="mouseSimulation">optional object mouseSimulation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions(object commandKeyHelp, object docNavigationKeys, object mouseSimulation)
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions", commandKeyHelp, docNavigationKeys, mouseSimulation);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandKeyHelp">optional object commandKeyHelp</param>
		/// <param name="docNavigationKeys">optional object docNavigationKeys</param>
		/// <param name="mouseSimulation">optional object mouseSimulation</param>
		/// <param name="demoGuidance">optional object demoGuidance</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions(object commandKeyHelp, object docNavigationKeys, object mouseSimulation, object demoGuidance)
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions", commandKeyHelp, docNavigationKeys, mouseSimulation, demoGuidance);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commandKeyHelp">optional object commandKeyHelp</param>
		/// <param name="docNavigationKeys">optional object docNavigationKeys</param>
		/// <param name="mouseSimulation">optional object mouseSimulation</param>
		/// <param name="demoGuidance">optional object demoGuidance</param>
		/// <param name="demoSpeed">optional object demoSpeed</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWPHelpOptions(object commandKeyHelp, object docNavigationKeys, object mouseSimulation, object demoGuidance, object demoSpeed)
		{
			 Factory.ExecuteMethod(this, "SetWPHelpOptions", new object[]{ commandKeyHelp, docNavigationKeys, mouseSimulation, demoGuidance, demoSpeed });
		}

		#endregion

		#pragma warning restore
	}
}
