using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IDialogSheet 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IDialogSheet : COMObject
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
                    _type = typeof(IDialogSheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDialogSheet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDialogSheet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDialogSheet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDialogSheet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDialogSheet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDialogSheet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDialogSheet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDialogSheet(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
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
		/// Get
		/// </summary>
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
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Next
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Next");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnDoubleClick
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDoubleClick");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDoubleClick", value);
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
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PageSetup PageSetup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PageSetup>(this, "PageSetup", NetOffice.ExcelApi.PageSetup.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Previous
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Previous");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectContents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectContents");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectDrawingObjects
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectDrawingObjects");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectionMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectionMode");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ProtectScenarios
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectScenarios");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSheetVisibility Visible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSheetVisibility>(this, "Visible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shapes Shapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shapes>(this, "Shapes", NetOffice.ExcelApi.Shapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableCalculation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableCalculation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableCalculation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayAutomaticPageBreaks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAutomaticPageBreaks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAutomaticPageBreaks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableAutoFilter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableAutoFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableAutoFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlEnableSelection EnableSelection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlEnableSelection>(this, "EnableSelection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EnableSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableOutlining
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableOutlining");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableOutlining", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnablePivotTable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnablePivotTable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnablePivotTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
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
		public string ScrollArea
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ScrollArea");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScrollArea", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.HPageBreaks HPageBreaks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.HPageBreaks>(this, "HPageBreaks", NetOffice.ExcelApi.HPageBreaks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.VPageBreaks VPageBreaks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.VPageBreaks>(this, "VPageBreaks", NetOffice.ExcelApi.VPageBreaks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.QueryTables QueryTables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.QueryTables>(this, "QueryTables", NetOffice.ExcelApi.QueryTables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayPageBreaks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPageBreaks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPageBreaks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Comments Comments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Comments>(this, "Comments", NetOffice.ExcelApi.Comments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Hyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Hyperlinks>(this, "Hyperlinks", NetOffice.ExcelApi.Hyperlinks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 _DisplayRightToLeft
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "_DisplayRightToLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_DisplayRightToLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.AutoFilter AutoFilter
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AutoFilter>(this, "AutoFilter", NetOffice.ExcelApi.AutoFilter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DisplayRightToLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayRightToLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayRightToLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Scripts>(this, "Scripts", NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DefaultButton
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultButton");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.DialogFrame DialogFrame
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DialogFrame>(this, "DialogFrame", NetOffice.ExcelApi.DialogFrame.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Focus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Focus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Focus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Tab Tab
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Tab>(this, "Tab", NetOffice.ExcelApi.Tab.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.MsoEnvelope MailEnvelope
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoEnvelope>(this, "MailEnvelope", NetOffice.OfficeApi.MsoEnvelope.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.CustomProperties CustomProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CustomProperties>(this, "CustomProperties", NetOffice.ExcelApi.CustomProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.SmartTags SmartTags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SmartTags>(this, "SmartTags", NetOffice.ExcelApi.SmartTags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Protection Protection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Protection>(this, "Protection", NetOffice.ExcelApi.Protection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool EnableFormatConditionsCalculation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableFormatConditionsCalculation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableFormatConditionsCalculation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Sort Sort
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sort>(this, "Sort", NetOffice.ExcelApi.Sort.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 PrintedCommentPages
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PrintedCommentPages");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Activate()
		{
			return Factory.ExecuteInt32MethodGet(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Copy(object before, object after)
		{
			return Factory.ExecuteInt32MethodGet(this, "Copy", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Copy()
		{
			return Factory.ExecuteInt32MethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Copy(object before)
		{
			return Factory.ExecuteInt32MethodGet(this, "Copy", before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Delete()
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Move(object before, object after)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Move()
		{
			return Factory.ExecuteInt32MethodGet(this, "Move");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Move(object before)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", before);
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
		public Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
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
		public Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _PrintOut()
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _PrintOut(object from)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _PrintOut(object from, object to)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", from, to);
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
		public Int32 _PrintOut(object from, object to, object copies)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", from, to, copies);
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
		public Int32 _PrintOut(object from, object to, object copies, object preview)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", from, to, copies, preview);
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
		public Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter });
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
		public Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="enableChanges">optional object enableChanges</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintPreview(object enableChanges)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintPreview", enableChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintPreview()
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object allowDeletingRows</param>
		/// <param name="allowSorting">optional object allowSorting</param>
		/// <param name="allowFiltering">optional object allowFiltering</param>
		/// <param name="allowUsingPivotTables">optional object allowUsingPivotTables</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering, object allowUsingPivotTables)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTables });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Protect()
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Protect(object password)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", password, drawingObjects);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", password, drawingObjects, contents);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", password, drawingObjects, contents, scenarios);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object allowDeletingRows</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object allowDeletingRows</param>
		/// <param name="allowSorting">optional object allowSorting</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object allowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object allowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object allowDeletingRows</param>
		/// <param name="allowSorting">optional object allowSorting</param>
		/// <param name="allowFiltering">optional object allowFiltering</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering)
		{
			return Factory.ExecuteInt32MethodGet(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="local">optional object local</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout, object local)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", filename, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", filename, fileFormat, password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", filename, fileFormat, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
		{
			return Factory.ExecuteInt32MethodGet(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Select(object replace)
		{
			return Factory.ExecuteInt32MethodGet(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Select()
		{
			return Factory.ExecuteInt32MethodGet(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Unprotect(object password)
		{
			return Factory.ExecuteInt32MethodGet(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Unprotect()
		{
			return Factory.ExecuteInt32MethodGet(this, "Unprotect");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy29()
		{
			 Factory.ExecuteMethod(this, "_Dummy29");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Arcs(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Arcs", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Arcs()
		{
			return Factory.ExecuteVariantMethodGet(this, "Arcs");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy31()
		{
			 Factory.ExecuteMethod(this, "_Dummy31");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy32()
		{
			 Factory.ExecuteMethod(this, "_Dummy32");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Buttons(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Buttons", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Buttons()
		{
			return Factory.ExecuteVariantMethodGet(this, "Buttons");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy34()
		{
			 Factory.ExecuteMethod(this, "_Dummy34");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy36()
		{
			 Factory.ExecuteMethod(this, "_Dummy36");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ChartObjects(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ChartObjects", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ChartObjects()
		{
			return Factory.ExecuteVariantMethodGet(this, "ChartObjects");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckBoxes(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckBoxes", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CheckBoxes()
		{
			return Factory.ExecuteVariantMethodGet(this, "CheckBoxes");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CheckSpelling()
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CheckSpelling(object customDictionary)
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy40()
		{
			 Factory.ExecuteMethod(this, "_Dummy40");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy41()
		{
			 Factory.ExecuteMethod(this, "_Dummy41");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy42()
		{
			 Factory.ExecuteMethod(this, "_Dummy42");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy43()
		{
			 Factory.ExecuteMethod(this, "_Dummy43");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy44()
		{
			 Factory.ExecuteMethod(this, "_Dummy44");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy45()
		{
			 Factory.ExecuteMethod(this, "_Dummy45");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Drawings(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Drawings", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Drawings()
		{
			return Factory.ExecuteVariantMethodGet(this, "Drawings");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DrawingObjects(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "DrawingObjects", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DrawingObjects()
		{
			return Factory.ExecuteVariantMethodGet(this, "DrawingObjects");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DropDowns(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "DropDowns", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object DropDowns()
		{
			return Factory.ExecuteVariantMethodGet(this, "DropDowns");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">object name</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Evaluate(object name)
		{
			return Factory.ExecuteVariantMethodGet(this, "Evaluate", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object _Evaluate(object name)
		{
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy56()
		{
			 Factory.ExecuteMethod(this, "_Dummy56");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 ResetAllPageBreaks()
		{
			return Factory.ExecuteInt32MethodGet(this, "ResetAllPageBreaks");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object GroupBoxes(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "GroupBoxes", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object GroupBoxes()
		{
			return Factory.ExecuteVariantMethodGet(this, "GroupBoxes");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object GroupObjects(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "GroupObjects", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object GroupObjects()
		{
			return Factory.ExecuteVariantMethodGet(this, "GroupObjects");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Labels(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Labels", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Labels()
		{
			return Factory.ExecuteVariantMethodGet(this, "Labels");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Lines(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Lines", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Lines()
		{
			return Factory.ExecuteVariantMethodGet(this, "Lines");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ListBoxes(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ListBoxes", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ListBoxes()
		{
			return Factory.ExecuteVariantMethodGet(this, "ListBoxes");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OLEObjects(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "OLEObjects", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OLEObjects()
		{
			return Factory.ExecuteVariantMethodGet(this, "OLEObjects");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy65()
		{
			 Factory.ExecuteMethod(this, "_Dummy65");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy66()
		{
			 Factory.ExecuteMethod(this, "_Dummy66");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy67()
		{
			 Factory.ExecuteMethod(this, "_Dummy67");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OptionButtons(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "OptionButtons", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OptionButtons()
		{
			return Factory.ExecuteVariantMethodGet(this, "OptionButtons");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy69()
		{
			 Factory.ExecuteMethod(this, "_Dummy69");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Ovals(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Ovals", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Ovals()
		{
			return Factory.ExecuteVariantMethodGet(this, "Ovals");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="destination">optional object destination</param>
		/// <param name="link">optional object link</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Paste(object destination, object link)
		{
			return Factory.ExecuteInt32MethodGet(this, "Paste", destination, link);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Paste()
		{
			return Factory.ExecuteInt32MethodGet(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Paste(object destination)
		{
			return Factory.ExecuteInt32MethodGet(this, "Paste", destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="noHTMLFormatting">optional object noHTMLFormatting</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object noHTMLFormatting)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex, iconLabel, noHTMLFormatting });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial()
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format, object link)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", format, link);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format, object link, object displayAsIcon)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", format, link, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", format, link, displayAsIcon, iconFileName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Pictures(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Pictures", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Pictures()
		{
			return Factory.ExecuteVariantMethodGet(this, "Pictures");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy74()
		{
			 Factory.ExecuteMethod(this, "_Dummy74");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy75()
		{
			 Factory.ExecuteMethod(this, "_Dummy75");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy76()
		{
			 Factory.ExecuteMethod(this, "_Dummy76");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Rectangles(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Rectangles", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Rectangles()
		{
			return Factory.ExecuteVariantMethodGet(this, "Rectangles");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy78()
		{
			 Factory.ExecuteMethod(this, "_Dummy78");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy79()
		{
			 Factory.ExecuteMethod(this, "_Dummy79");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ScrollBars(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ScrollBars", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ScrollBars()
		{
			return Factory.ExecuteVariantMethodGet(this, "ScrollBars");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy82()
		{
			 Factory.ExecuteMethod(this, "_Dummy82");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy83()
		{
			 Factory.ExecuteMethod(this, "_Dummy83");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Spinners(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Spinners", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Spinners()
		{
			return Factory.ExecuteVariantMethodGet(this, "Spinners");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy85()
		{
			 Factory.ExecuteMethod(this, "_Dummy85");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy86()
		{
			 Factory.ExecuteMethod(this, "_Dummy86");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextBoxes(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "TextBoxes", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextBoxes()
		{
			return Factory.ExecuteVariantMethodGet(this, "TextBoxes");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy88()
		{
			 Factory.ExecuteMethod(this, "_Dummy88");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy89()
		{
			 Factory.ExecuteMethod(this, "_Dummy89");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy90()
		{
			 Factory.ExecuteMethod(this, "_Dummy90");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 ClearCircles()
		{
			return Factory.ExecuteInt32MethodGet(this, "ClearCircles");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 CircleInvalid()
		{
			return Factory.ExecuteInt32MethodGet(this, "CircleInvalid");
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
		/// <param name="prToFileName">optional object prToFileName</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut()
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to, object copies)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to, object copies, object preview)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter });
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			return Factory.ExecuteInt32MethodGet(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		/// <param name="ignoreFinalYaa">optional object ignoreFinalYaa</param>
		/// <param name="spellScript">optional object spellScript</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang, object ignoreFinalYaa, object spellScript)
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa, spellScript });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling()
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling(object customDictionary)
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		/// <param name="ignoreFinalYaa">optional object ignoreFinalYaa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang, object ignoreFinalYaa)
		{
			return Factory.ExecuteInt32MethodGet(this, "_CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditBoxes(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "EditBoxes", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object EditBoxes()
		{
			return Factory.ExecuteVariantMethodGet(this, "EditBoxes");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cancel">optional object cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Hide(object cancel)
		{
			return Factory.ExecuteBoolMethodGet(this, "Hide", cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Hide()
		{
			return Factory.ExecuteBoolMethodGet(this, "Hide");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Show()
		{
			return Factory.ExecuteBoolMethodGet(this, "Show");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			return Factory.ExecuteInt32MethodGet(this, "_Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _Protect()
		{
			return Factory.ExecuteInt32MethodGet(this, "_Protect");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _Protect(object password)
		{
			return Factory.ExecuteInt32MethodGet(this, "_Protect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _Protect(object password, object drawingObjects)
		{
			return Factory.ExecuteInt32MethodGet(this, "_Protect", password, drawingObjects);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _Protect(object password, object drawingObjects, object contents)
		{
			return Factory.ExecuteInt32MethodGet(this, "_Protect", password, drawingObjects, contents);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			return Factory.ExecuteInt32MethodGet(this, "_Protect", password, drawingObjects, contents, scenarios);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", filename, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", filename, fileFormat, password);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", filename, fileFormat, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="textCodepage">optional object textCodepage</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
		{
			return Factory.ExecuteInt32MethodGet(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial()
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial(object format)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial(object format, object link)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial", format, link);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial(object format, object link, object displayAsIcon)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial", format, link, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial", format, link, displayAsIcon, iconFileName);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteInt32MethodGet(this, "_PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void _Dummy113()
		{
			 Factory.ExecuteMethod(this, "_Dummy113");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void _Dummy114()
		{
			 Factory.ExecuteMethod(this, "_Dummy114");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void _Dummy115()
		{
			 Factory.ExecuteMethod(this, "_Dummy115");
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
		public Int32 __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 __PrintOut()
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 __PrintOut(object from)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 __PrintOut(object from, object to)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", from, to);
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
		public Int32 __PrintOut(object from, object to, object copies)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", from, to, copies);
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
		public Int32 __PrintOut(object from, object to, object copies, object preview)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", from, to, copies, preview);
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
		public Int32 __PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter });
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
		public Int32 __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			return Factory.ExecuteInt32MethodGet(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
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
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename, quality);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename, quality, includeDocProperties);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="quality">optional object quality</param>
		/// <param name="includeDocProperties">optional object includeDocProperties</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
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
		public Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
		{
			return Factory.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[]{ type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish });
		}

		#endregion

		#pragma warning restore
	}
}
