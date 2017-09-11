using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// _Workbook
	/// </summary>
	[SyntaxBypass]
 	public class _Workbook_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Workbook_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Workbook_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Colors(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Colors", index);
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Colors(object index, object value)
		{
			Factory.ExecutePropertySet(this, "Colors", index, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Colors
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Colors")]
		public object Colors(object index)
		{
			return get_Colors(index);
		}

		#endregion

		#region Methods

		#endregion

	}

	/// <summary>
	/// DispatchInterface _Workbook 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Workbook : _Workbook_
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
                    _type = typeof(_Workbook);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Workbook(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Workbook(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Workbook(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835918.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840080.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198008.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool AcceptLabelsInFormulas
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AcceptLabelsInFormulas");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AcceptLabelsInFormulas", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834923.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart ActiveChart
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Chart>(this, "ActiveChart", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841181.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object ActiveSheet
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveSheet");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Author
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Author");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Author", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840067.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 AutoUpdateFrequency
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoUpdateFrequency");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoUpdateFrequency", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193298.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool AutoUpdateSaveChanges
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoUpdateSaveChanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoUpdateSaveChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821530.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 ChangeHistoryDuration
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ChangeHistoryDuration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ChangeHistoryDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197172.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object BuiltinDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "BuiltinDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821062.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Charts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Charts", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195162.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string CodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CodeName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _CodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_CodeName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_CodeName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Colors
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Colors");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Colors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835614.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Comments
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Comments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Comments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198339.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSaveConflictResolution ConflictResolution
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSaveConflictResolution>(this, "ConflictResolution");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ConflictResolution", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834401.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Container
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196337.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CreateBackup
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CreateBackup");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834990.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object CustomDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193264.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Date1904
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Date1904");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Date1904", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Sheets DialogSheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "DialogSheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834329.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects DisplayDrawingObjects
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects>(this, "DisplayDrawingObjects");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayDrawingObjects", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840717.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlFileFormat FileFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlFileFormat>(this, "FileFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834975.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasMailer
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMailer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMailer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840238.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool HasPassword
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPassword");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool HasRoutingSlip
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasRoutingSlip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasRoutingSlip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838249.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool IsAddin
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsAddin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsAddin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Keywords
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Keywords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Keywords", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837965.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Mailer Mailer
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Mailer>(this, "Mailer", NetOffice.ExcelApi.Mailer.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Sheets Modules
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Modules", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839882.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool MultiUserEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MultiUserEditing");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820899.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195422.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Names Names
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Names>(this, "Names", NetOffice.ExcelApi.Names.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnSave
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnSheetActivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnSheetActivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnSheetActivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnSheetDeactivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnSheetDeactivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnSheetDeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840974.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836500.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PersonalViewListSettings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PersonalViewListSettings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PersonalViewListSettings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822649.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PersonalViewPrintSettings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PersonalViewPrintSettings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PersonalViewPrintSettings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198189.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PrecisionAsDisplayed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrecisionAsDisplayed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrecisionAsDisplayed", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838601.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectStructure
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectStructure");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193864.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectWindows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectWindows");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840925.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ReadOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196964.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ReadOnlyRecommended
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnlyRecommended");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadOnlyRecommended", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834665.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 RevisionNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Routed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Routed");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.RoutingSlip RoutingSlip
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.RoutingSlip>(this, "RoutingSlip", NetOffice.ExcelApi.RoutingSlip.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196613.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Saved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Saved");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Saved", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840667.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SaveLinkValues
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveLinkValues");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveLinkValues", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197568.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Sheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Sheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839677.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ShowConflictHistory
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowConflictHistory");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowConflictHistory", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839039.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Styles Styles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Styles>(this, "Styles", NetOffice.ExcelApi.Styles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Subject
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Subject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Subject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193266.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool UpdateRemoteReferences
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateRemoteReferences");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateRemoteReferences", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UserControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UserControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193788.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object UserStatus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "UserStatus");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195531.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CustomViews CustomViews
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CustomViews>(this, "CustomViews", NetOffice.ExcelApi.CustomViews.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195152.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Windows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Windows>(this, "Windows", NetOffice.ExcelApi.Windows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835542.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Worksheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Worksheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836228.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool WriteReserved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WriteReserved");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840737.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string WriteReservedBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WriteReservedBy");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822819.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Excel4IntlMacroSheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Excel4IntlMacroSheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195645.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Sheets Excel4MacroSheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Excel4MacroSheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836472.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TemplateRemoveExtData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TemplateRemoveExtData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TemplateRemoveExtData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194254.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool HighlightChangesOnScreen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HighlightChangesOnScreen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HighlightChangesOnScreen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197016.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool KeepChangeHistory
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KeepChangeHistory");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeepChangeHistory", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834301.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ListChangesOnNewSheet
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ListChangesOnNewSheet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListChangesOnNewSheet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194737.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840963.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool IsInplace
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsInplace");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838208.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PublishObjects PublishObjects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PublishObjects>(this, "PublishObjects", NetOffice.ExcelApi.PublishObjects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834724.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.WebOptions WebOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WebOptions>(this, "WebOptions", NetOffice.ExcelApi.WebOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProject HTMLProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProject>(this, "HTMLProject", NetOffice.OfficeApi.HTMLProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839554.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnvelopeVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnvelopeVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnvelopeVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193512.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CalculationVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CalculationVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822659.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool VBASigned
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VBASigned");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool _ReadOnlyRecommended
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "_ReadOnlyRecommended");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196322.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool ShowPivotTableFieldList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPivotTableFieldList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPivotTableFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839021.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlUpdateLinks UpdateLinks
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlUpdateLinks>(this, "UpdateLinks");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "UpdateLinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193225.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EnableAutoRecover
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableAutoRecover");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableAutoRecover", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841017.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RemovePersonalInformation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RemovePersonalInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821089.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string FullNameURLEncoded
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullNameURLEncoded");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821529.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837767.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839579.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string PasswordEncryptionProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PasswordEncryptionProvider");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195464.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string PasswordEncryptionAlgorithm
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PasswordEncryptionAlgorithm");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195381.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 PasswordEncryptionKeyLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PasswordEncryptionKeyLength");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820819.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool PasswordEncryptionFileProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PasswordEncryptionFileProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SmartTagOptions SmartTagOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SmartTagOptions>(this, "SmartTagOptions", NetOffice.ExcelApi.SmartTagOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840697.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Permission Permission
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Permission>(this, "Permission", NetOffice.OfficeApi.Permission.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835236.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192923.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Sync Sync
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Sync>(this, "Sync", NetOffice.OfficeApi.Sync.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838260.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XmlNamespaces XmlNamespaces
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlNamespaces>(this, "XmlNamespaces", NetOffice.ExcelApi.XmlNamespaces.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838975.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XmlMaps XmlMaps
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlMaps>(this, "XmlMaps", NetOffice.ExcelApi.XmlMaps.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194561.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SmartDocument SmartDocument
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartDocument>(this, "SmartDocument", NetOffice.OfficeApi.SmartDocument.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838205.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837429.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool InactiveListBorderVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InactiveListBorderVisible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InactiveListBorderVisible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838435.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool DisplayInkComments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayInkComments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayInkComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837152.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.MetaProperties ContentTypeProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MetaProperties>(this, "ContentTypeProperties", NetOffice.OfficeApi.MetaProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836773.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Connections Connections
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Connections>(this, "Connections", NetOffice.ExcelApi.Connections.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838073.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.SignatureSet Signatures
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSet>(this, "Signatures", NetOffice.OfficeApi.SignatureSet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194489.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.ServerPolicy ServerPolicy
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ServerPolicy>(this, "ServerPolicy", NetOffice.OfficeApi.ServerPolicy.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195426.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentInspectors>(this, "DocumentInspectors", NetOffice.OfficeApi.DocumentInspectors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195818.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ServerViewableItems ServerViewableItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ServerViewableItems>(this, "ServerViewableItems", NetOffice.ExcelApi.ServerViewableItems.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837756.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.TableStyles TableStyles
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TableStyles>(this, "TableStyles", NetOffice.ExcelApi.TableStyles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195934.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultTableStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultTableStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultTableStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835624.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultPivotTableStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultPivotTableStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultPivotTableStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836165.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool CheckCompatibility
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CheckCompatibility");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CheckCompatibility", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838063.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool HasVBProject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasVBProject");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838448.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLParts>(this, "CustomXMLParts", NetOffice.OfficeApi.CustomXMLParts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820907.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196847.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Research Research
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Research>(this, "Research", NetOffice.ExcelApi.Research.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194072.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.OfficeTheme Theme
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeTheme>(this, "Theme", NetOffice.OfficeApi.OfficeTheme.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834991.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool Excel8CompatibilityMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Excel8CompatibilityMode");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837960.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ConnectionsDisabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConnectionsDisabled");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835280.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowPivotChartActiveFields
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPivotChartActiveFields");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPivotChartActiveFields", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839003.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.IconSets IconSets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.IconSets>(this, "IconSets", NetOffice.ExcelApi.IconSets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194147.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839440.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823189.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ForceFullCalculation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ForceFullCalculation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForceFullCalculation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194925.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerCaches SlicerCaches
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerCaches>(this, "SlicerCaches", NetOffice.ExcelApi.SlicerCaches.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839464.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer ActiveSlicer
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicer>(this, "ActiveSlicer", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193862.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultSlicerStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultSlicerStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultSlicerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838425.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 AccuracyVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AccuracyVersion");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AccuracyVersion", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229542.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool CaseSensitive
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CaseSensitive");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231362.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool UseWholeCellCriteria
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseWholeCellCriteria");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230772.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool UseWildcards
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseWildcards");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231277.aspx </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		public object PivotTables
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "PivotTables");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228926.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Model Model
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Model>(this, "Model", NetOffice.ExcelApi.Model.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227452.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
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
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230214.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DefaultTimelineStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultTimelineStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultTimelineStyle", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821837.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
		/// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
		/// <param name="writePassword">optional object writePassword</param>
		/// <param name="notify">optional object notify</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword, object notify)
		{
			 Factory.ExecuteMethod(this, "ChangeFileAccess", mode, writePassword, notify);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
		/// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode)
		{
			 Factory.ExecuteMethod(this, "ChangeFileAccess", mode);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
		/// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
		/// <param name="writePassword">optional object writePassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword)
		{
			 Factory.ExecuteMethod(this, "ChangeFileAccess", mode, writePassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="newName">string newName</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlLinkType Type = 1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ChangeLink(string name, string newName, object type)
		{
			 Factory.ExecuteMethod(this, "ChangeLink", name, newName, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="newName">string newName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ChangeLink(string name, string newName)
		{
			 Factory.ExecuteMethod(this, "ChangeLink", name, newName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="routeWorkbook">optional object routeWorkbook</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object filename, object routeWorkbook)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges, filename, routeWorkbook);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object filename)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839565.aspx </remarks>
		/// <param name="numberFormat">string numberFormat</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void DeleteNumberFormat(string numberFormat)
		{
			 Factory.ExecuteMethod(this, "DeleteNumberFormat", numberFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836762.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ExclusiveAccess()
		{
			return Factory.ExecuteBoolMethodGet(this, "ExclusiveAccess");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836208.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ForwardMailer()
		{
			 Factory.ExecuteMethod(this, "ForwardMailer");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
		/// <param name="type">optional object type</param>
		/// <param name="editionRef">optional object editionRef</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type, object editionRef)
		{
			return Factory.ExecuteVariantMethodGet(this, "LinkInfo", name, linkInfo, type, editionRef);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo)
		{
			return Factory.ExecuteVariantMethodGet(this, "LinkInfo", name, linkInfo);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "LinkInfo", name, linkInfo, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx </remarks>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object LinkSources(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "LinkSources", type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object LinkSources()
		{
			return Factory.ExecuteVariantMethodGet(this, "LinkSources");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196693.aspx </remarks>
		/// <param name="filename">object filename</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void MergeWorkbook(object filename)
		{
			 Factory.ExecuteMethod(this, "MergeWorkbook", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838378.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Window NewWindow()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Window>(this, "NewWindow", NetOffice.ExcelApi.Window.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenLinks(string name, object readOnly, object type)
		{
			 Factory.ExecuteMethod(this, "OpenLinks", name, readOnly, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenLinks(string name)
		{
			 Factory.ExecuteMethod(this, "OpenLinks", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenLinks(string name, object readOnly)
		{
			 Factory.ExecuteMethod(this, "OpenLinks", name, readOnly);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193549.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCaches PivotCaches()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCaches>(this, "PivotCaches", NetOffice.ExcelApi.PivotCaches.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx </remarks>
		/// <param name="destName">optional object destName</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Post(object destName)
		{
			 Factory.ExecuteMethod(this, "Post", destName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Post()
		{
			 Factory.ExecuteMethod(this, "Post");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut()
		{
			 Factory.ExecuteMethod(this, "_PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx </remarks>
		/// <param name="enableChanges">optional object enableChanges</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview(object enableChanges)
		{
			 Factory.ExecuteMethod(this, "PrintPreview", enableChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			 Factory.ExecuteMethod(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="structure">optional object structure</param>
		/// <param name="windows">optional object windows</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object structure, object windows)
		{
			 Factory.ExecuteMethod(this, "Protect", password, structure, windows);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect()
		{
			 Factory.ExecuteMethod(this, "Protect");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password)
		{
			 Factory.ExecuteMethod(this, "Protect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="structure">optional object structure</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object structure)
		{
			 Factory.ExecuteMethod(this, "Protect", password, structure);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="sharingPassword">optional object sharingPassword</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", new object[]{ filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="sharingPassword">optional object sharingPassword</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", new object[]{ filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword, fileFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing()
		{
			 Factory.ExecuteMethod(this, "ProtectSharing");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", filename, password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", filename, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", filename, password, writeResPassword, readOnlyRecommended);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			 Factory.ExecuteMethod(this, "ProtectSharing", new object[]{ filename, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838648.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RefreshAll()
		{
			 Factory.ExecuteMethod(this, "RefreshAll");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820902.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Reply()
		{
			 Factory.ExecuteMethod(this, "Reply");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838788.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ReplyAll()
		{
			 Factory.ExecuteMethod(this, "ReplyAll");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840747.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RemoveUser(Int32 index)
		{
			 Factory.ExecuteMethod(this, "RemoveUser", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Route()
		{
			 Factory.ExecuteMethod(this, "Route");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835203.aspx </remarks>
		/// <param name="which">NetOffice.ExcelApi.Enums.XlRunAutoMacro which</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RunAutoMacros(NetOffice.ExcelApi.Enums.XlRunAutoMacro which)
		{
			 Factory.ExecuteMethod(this, "RunAutoMacros", which);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197585.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="local">optional object local</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout, object local)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs()
		{
			 Factory.ExecuteMethod(this, "SaveAs");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, fileFormat, password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, fileFormat, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx </remarks>
		/// <param name="filename">optional object filename</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveCopyAs(object filename)
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveCopyAs()
		{
			 Factory.ExecuteMethod(this, "SaveCopyAs");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
		/// <param name="recipients">object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="returnReceipt">optional object returnReceipt</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SendMail(object recipients, object subject, object returnReceipt)
		{
			 Factory.ExecuteMethod(this, "SendMail", recipients, subject, returnReceipt);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
		/// <param name="recipients">object recipients</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SendMail(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendMail", recipients);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
		/// <param name="recipients">object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SendMail(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendMail", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="priority">optional NetOffice.ExcelApi.Enums.XlPriority Priority = -4143</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat, object priority)
		{
			 Factory.ExecuteMethod(this, "SendMailer", fileFormat, priority);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SendMailer()
		{
			 Factory.ExecuteMethod(this, "SendMailer");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SendMailer(object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SendMailer", fileFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="procedure">optional object procedure</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SetLinkOnData(string name, object procedure)
		{
			 Factory.ExecuteMethod(this, "SetLinkOnData", name, procedure);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SetLinkOnData(string name)
		{
			 Factory.ExecuteMethod(this, "SetLinkOnData", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx </remarks>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Unprotect(object password)
		{
			 Factory.ExecuteMethod(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Unprotect()
		{
			 Factory.ExecuteMethod(this, "Unprotect");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx </remarks>
		/// <param name="sharingPassword">optional object sharingPassword</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UnprotectSharing(object sharingPassword)
		{
			 Factory.ExecuteMethod(this, "UnprotectSharing", sharingPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UnprotectSharing()
		{
			 Factory.ExecuteMethod(this, "UnprotectSharing");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840979.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UpdateFromFile()
		{
			 Factory.ExecuteMethod(this, "UpdateFromFile");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UpdateLink(object name, object type)
		{
			 Factory.ExecuteMethod(this, "UpdateLink", name, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UpdateLink()
		{
			 Factory.ExecuteMethod(this, "UpdateLink");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void UpdateLink(object name)
		{
			 Factory.ExecuteMethod(this, "UpdateLink", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
		/// <param name="when">optional object when</param>
		/// <param name="who">optional object who</param>
		/// <param name="where">optional object where</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions(object when, object who, object where)
		{
			 Factory.ExecuteMethod(this, "HighlightChangesOptions", when, who, where);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions()
		{
			 Factory.ExecuteMethod(this, "HighlightChangesOptions");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
		/// <param name="when">optional object when</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions(object when)
		{
			 Factory.ExecuteMethod(this, "HighlightChangesOptions", when);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
		/// <param name="when">optional object when</param>
		/// <param name="who">optional object who</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void HighlightChangesOptions(object when, object who)
		{
			 Factory.ExecuteMethod(this, "HighlightChangesOptions", when, who);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx </remarks>
		/// <param name="days">Int32 days</param>
		/// <param name="sharingPassword">optional object sharingPassword</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PurgeChangeHistoryNow(Int32 days, object sharingPassword)
		{
			 Factory.ExecuteMethod(this, "PurgeChangeHistoryNow", days, sharingPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx </remarks>
		/// <param name="days">Int32 days</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PurgeChangeHistoryNow(Int32 days)
		{
			 Factory.ExecuteMethod(this, "PurgeChangeHistoryNow", days);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
		/// <param name="when">optional object when</param>
		/// <param name="who">optional object who</param>
		/// <param name="where">optional object where</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges(object when, object who, object where)
		{
			 Factory.ExecuteMethod(this, "AcceptAllChanges", when, who, where);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges()
		{
			 Factory.ExecuteMethod(this, "AcceptAllChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
		/// <param name="when">optional object when</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges(object when)
		{
			 Factory.ExecuteMethod(this, "AcceptAllChanges", when);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
		/// <param name="when">optional object when</param>
		/// <param name="who">optional object who</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AcceptAllChanges(object when, object who)
		{
			 Factory.ExecuteMethod(this, "AcceptAllChanges", when, who);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
		/// <param name="when">optional object when</param>
		/// <param name="who">optional object who</param>
		/// <param name="where">optional object where</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges(object when, object who, object where)
		{
			 Factory.ExecuteMethod(this, "RejectAllChanges", when, who, where);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges()
		{
			 Factory.ExecuteMethod(this, "RejectAllChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
		/// <param name="when">optional object when</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges(object when)
		{
			 Factory.ExecuteMethod(this, "RejectAllChanges", when);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
		/// <param name="when">optional object when</param>
		/// <param name="who">optional object who</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void RejectAllChanges(object when, object who)
		{
			 Factory.ExecuteMethod(this, "RejectAllChanges", when, who);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
		/// <param name="readData">optional object readData</param>
		/// <param name="connection">optional object connection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard()
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", sourceType, sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", sourceType, sourceData, tableDestination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", sourceType, sourceData, tableDestination, tableName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		/// <param name="reserved">optional object reserved</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="optimizeCache">optional object optimizeCache</param>
		/// <param name="pageFieldOrder">optional object pageFieldOrder</param>
		/// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
		/// <param name="readData">optional object readData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData)
		{
			 Factory.ExecuteMethod(this, "PivotTableWizard", new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194697.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ResetColors()
		{
			 Factory.ExecuteMethod(this, "ResetColors");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		/// <param name="headerInfo">optional object headerInfo</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional object subAddress</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
		{
			 Factory.ExecuteMethod(this, "FollowHyperlink", new object[]{ address, subAddress, newWindow, addHistory, extraInfo, method });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194282.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void AddToFavorites()
		{
			 Factory.ExecuteMethod(this, "AddToFavorites");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName, object ignorePrintAreas)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, ignorePrintAreas });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195831.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void WebPagePreview()
		{
			 Factory.ExecuteMethod(this, "WebPagePreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839234.aspx </remarks>
		/// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding encoding</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding)
		{
			 Factory.ExecuteMethod(this, "ReloadAs", encoding);
		}

		/// <summary>
		/// SupportByVersion Excel 9
		/// </summary>
		/// <param name="unused">Int32 unused</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9)]
		public void Dummy1(Int32 unused)
		{
			 Factory.ExecuteMethod(this, "Dummy1", unused);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void sblt(string s)
		{
			 Factory.ExecuteMethod(this, "sblt", s);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="structure">optional object structure</param>
		/// <param name="windows">optional object windows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object structure, object windows)
		{
			 Factory.ExecuteMethod(this, "_Protect", password, structure, windows);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _Protect()
		{
			 Factory.ExecuteMethod(this, "_Protect");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password)
		{
			 Factory.ExecuteMethod(this, "_Protect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="structure">optional object structure</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object structure)
		{
			 Factory.ExecuteMethod(this, "_Protect", password, structure);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs()
		{
			 Factory.ExecuteMethod(this, "_SaveAs");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename, fileFormat, password);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename, fileFormat, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
		/// <param name="conflictResolution">optional object conflictResolution</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="calcid">Int32 calcid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Dummy17(Int32 calcid)
		{
			 Factory.ExecuteMethod(this, "Dummy17", calcid);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194915.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlLinkType type</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void BreakLink(string name, NetOffice.ExcelApi.Enums.XlLinkType type)
		{
			 Factory.ExecuteMethod(this, "BreakLink", name, type);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Dummy16()
		{
			 Factory.ExecuteMethod(this, "Dummy16");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void CheckIn()
		{
			 Factory.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194456.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool CanCheckIn()
		{
			return Factory.ExecuteBoolMethodGet(this, "CanCheckIn");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		/// <param name="includeAttachment">optional object includeAttachment</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage, includeAttachment);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SendForReview()
		{
			 Factory.ExecuteMethod(this, "SendForReview");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SendForReview(object recipients, object subject, object showMessage)
		{
			 Factory.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx </remarks>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void ReplyWithChanges(object showMessage)
		{
			 Factory.ExecuteMethod(this, "ReplyWithChanges", showMessage);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void ReplyWithChanges()
		{
			 Factory.ExecuteMethod(this, "ReplyWithChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839207.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void EndReview()
		{
			 Factory.ExecuteMethod(this, "EndReview");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
		/// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">optional object passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">optional object passwordEncryptionFileProperties</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions()
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
		/// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
		/// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
		/// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">optional object passwordEncryptionKeyLength</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength)
		{
			 Factory.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void RecheckSmartTags()
		{
			 Factory.ExecuteMethod(this, "RecheckSmartTags");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		/// <param name="showMessage">optional object showMessage</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject, object showMessage)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject, showMessage);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet()
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
		/// <param name="recipients">optional object recipients</param>
		/// <param name="subject">optional object subject</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void SendFaxOverInternet(object recipients, object subject)
		{
			 Factory.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
		/// <param name="url">string url</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
		/// <param name="overwrite">optional object overwrite</param>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination)
		{
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);            
            object[] paramsArray = Invoker.ValidateParamsArray(url, new object(), overwrite, destination);
            object returnItem = Invoker.MethodReturn(this, "XmlImport", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                importMap = new NetOffice.ExcelApi.XmlMap(this, paramsArray[1]);
            else
                importMap = null;
            int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
		/// <param name="url">string url</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			object[] paramsArray = Invoker.ValidateParamsArray(url, new object());
			object returnItem = Invoker.MethodReturn(this, "XmlImport", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                importMap = new NetOffice.ExcelApi.XmlMap(this, paramsArray[1]);
            else
                importMap = null;
            int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
		/// <param name="url">string url</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
		/// <param name="overwrite">optional object overwrite</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			object[] paramsArray = Invoker.ValidateParamsArray(url, new object(), overwrite);
			object returnItem = Invoker.MethodReturn(this, "XmlImport", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                importMap = new NetOffice.ExcelApi.XmlMap(this, paramsArray[1]);
            else
                importMap = null;
            int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
		/// <param name="data">string data</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
		/// <param name="overwrite">optional object overwrite</param>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);			
			object[] paramsArray = Invoker.ValidateParamsArray(data, new object(), overwrite, destination);
			object returnItem = Invoker.MethodReturn(this, "XmlImportXml", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                importMap = new NetOffice.ExcelApi.XmlMap(this, paramsArray[1]);
            else
                importMap = null;
            int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
		/// <param name="data">string data</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			object[] paramsArray = Invoker.ValidateParamsArray(data, new object());
			object returnItem = Invoker.MethodReturn(this, "XmlImportXml", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                importMap = new NetOffice.ExcelApi.XmlMap(this, paramsArray[1]);
            else
                importMap = null;
            int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
		/// <param name="data">string data</param>
		/// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
		/// <param name="overwrite">optional object overwrite</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);			 
			object[] paramsArray = Invoker.ValidateParamsArray(data, new object(), overwrite);
			object returnItem = Invoker.MethodReturn(this, "XmlImportXml", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                importMap = new NetOffice.ExcelApi.XmlMap(this, paramsArray[1]);
            else
                importMap = null;
            int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.ExcelApi.Enums.XlXmlImportResult)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834616.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void SaveAsXMLData(string filename, NetOffice.ExcelApi.XmlMap map)
		{
			 Factory.ExecuteMethod(this, "SaveAsXMLData", filename, map);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196845.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void ToggleFormsDesign()
		{
			 Factory.ExecuteMethod(this, "ToggleFormsDesign");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut()
		{
			 Factory.ExecuteMethod(this, "__PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="sharingPassword">optional object sharingPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword)
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing", new object[]{ filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing()
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename)
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password)
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing", filename, password);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword)
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing", filename, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing", filename, password, writeResPassword, readOnlyRecommended);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">optional object filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			 Factory.ExecuteMethod(this, "_ProtectSharing", new object[]{ filename, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840327.aspx </remarks>
		/// <param name="removeDocInfoType">NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void RemoveDocumentInformation(NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType)
		{
			 Factory.ExecuteMethod(this, "RemoveDocumentInformation", removeDocInfoType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic, versionType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CheckInWithVersion()
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CheckInWithVersion(object saveChanges, object comments, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838567.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void LockServerFile()
		{
			 Factory.ExecuteMethod(this, "LockServerFile");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835507.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTasks>(this, "GetWorkflowTasks", NetOffice.OfficeApi.WorkflowTasks.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837818.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTemplates>(this, "GetWorkflowTemplates", NetOffice.OfficeApi.WorkflowTemplates.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194014.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ApplyTheme(string filename)
		{
			 Factory.ExecuteMethod(this, "ApplyTheme", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820742.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void EnableConnections()
		{
			 Factory.ExecuteMethod(this, "EnableConnections");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="openAfterPublish">optional object openAfterPublish</param>
		/// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type, filename, quality);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", type, filename, quality, includeDocProperties);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="openAfterPublish">optional object openAfterPublish</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 14,15,16)]
		public void Dummy26()
		{
			 Factory.ExecuteMethod(this, "Dummy26");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 14,15,16)]
		public void Dummy27()
		{
			 Factory.ExecuteMethod(this, "Dummy27");
		}

		#endregion

		#pragma warning restore
	}
}
