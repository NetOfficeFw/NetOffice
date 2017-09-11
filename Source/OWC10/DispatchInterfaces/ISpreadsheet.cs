using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ISpreadsheet 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class ISpreadsheet : COMObject
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
                    _type = typeof(ISpreadsheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISpreadsheet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISpreadsheet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpreadsheet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpreadsheet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpreadsheet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpreadsheet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpreadsheet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpreadsheet(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range ActiveCell
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "ActiveCell");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet ActiveSheet
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "ActiveSheet", NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Window ActiveWindow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Window>(this, "ActiveWindow", NetOffice.OWC10Api.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Workbook ActiveWorkbook
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Workbook>(this, "ActiveWorkbook", NetOffice.OWC10Api.Workbook.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowPropertyToolbox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPropertyToolbox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AutoFit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Build
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Build");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string BuildNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlCalculation Calculation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlCalculation>(this, "Calculation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Calculation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 CalculationVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CalculationVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool CanUndo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanUndo");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Cells
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Cells");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Columns
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Columns");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.OCCommands Commands
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OCCommands>(this, "Commands", NetOffice.OWC10Api.OCCommands.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Constants
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Constants");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CSVData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CSVData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CSVData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string CSVURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CSVURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CSVURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataMember
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.MSDATASRCApi.DataSource DataSource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSDATASRCApi.DataSource>(this, "DataSource", NetOffice.MSDATASRCApi.DataSource.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DesignMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DesignMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Dirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayBranding
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayBranding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayBranding", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayColumnHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayColumnHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayColumnHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayDesignTimeUI
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayDesignTimeUI");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayDesignTimeUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayHorizontalScrollBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayHorizontalScrollBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayHorizontalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayOfficeLogo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayOfficeLogo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayOfficeLogo", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayPropertyToolbox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPropertyToolbox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayRowHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayRowHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayRowHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayTitleBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayTitleBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayTitleBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayToolbar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayToolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayVerticalScrollBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayVerticalScrollBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayVerticalScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayWorkbookTabs
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayWorkbookTabs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayWorkbookTabs", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool EnableEvents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableEvents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableEvents", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool EnableUndo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableUndo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableUndo", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string HTMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HTMLData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HTMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string HTMLURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HTMLURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HTMLURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 InstanceID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InstanceID");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_International(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "International", index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_International
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1), Redirect("get_International")]
		public object International(object index)
		{
			return get_International(index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.OWCLanguageSettings LanguageSettings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OWCLanguageSettings>(this, "LanguageSettings", NetOffice.OWC10Api.OWCLanguageSettings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object MaxHeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MaxHeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MaxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object MaxWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MaxWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MaxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MajorVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string MinorVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool MoveAfterReturn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MoveAfterReturn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MoveAfterReturn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlDirection MoveAfterReturnDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlDirection>(this, "MoveAfterReturnDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MoveAfterReturnDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Names Names
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Names>(this, "Names", NetOffice.OWC10Api.Names.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1, object cell2)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public NetOffice.OWC10Api._Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Range", NetOffice.OWC10Api._Range.LateBindingApiWrapperType, cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Range")]
		public NetOffice.OWC10Api._Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RevisionNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool RightToLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RightToLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightToLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Rows
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Rows");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Selection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Selection");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Sheets Sheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Sheets>(this, "Sheets", NetOffice.OWC10Api.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.TitleBar TitleBar
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.TitleBar>(this, "TitleBar", NetOffice.OWC10Api.TitleBar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.MSComctlLibApi.IToolbar Toolbar
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IToolbar>(this, "Toolbar");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Value");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ViewableRange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ViewableRange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewableRange", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ViewOnlyMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewOnlyMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Windows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Windows>(this, "Windows", NetOffice.OWC10Api.Windows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Workbooks Workbooks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Workbooks>(this, "Workbooks", NetOffice.OWC10Api.Workbooks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheets Worksheets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheets>(this, "Worksheets", NetOffice.OWC10Api.Worksheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string XMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string XMLURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLURL", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="addIn">object addIn</param>
		[SupportByVersion("OWC10", 1)]
		public void AddIn(object addIn)
		{
			 Factory.ExecuteMethod(this, "AddIn", addIn);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void BeginUndo()
		{
			 Factory.ExecuteMethod(this, "BeginUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Calculate()
		{
			 Factory.ExecuteMethod(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void CalculateFull()
		{
			 Factory.ExecuteMethod(this, "CalculateFull");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cancel">optional bool Cancel = false</param>
		[SupportByVersion("OWC10", 1)]
		public void EndUndo(object cancel)
		{
			 Factory.ExecuteMethod(this, "EndUndo", cancel);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void EndUndo()
		{
			 Factory.ExecuteMethod(this, "EndUndo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="expression">object expression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public object _Evaluate(object expression)
		{
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", expression);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="expression">object expression</param>
		[SupportByVersion("OWC10", 1)]
		public object Evaluate(object expression)
		{
			return Factory.ExecuteVariantMethodGet(this, "Evaluate", expression);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.SheetExportActionEnum Action = 1</param>
		/// <param name="format">optional NetOffice.OWC10Api.Enums.SheetExportFormat Format = 0</param>
		[SupportByVersion("OWC10", 1)]
		public void Export(object filename, object action, object format)
		{
			 Factory.ExecuteMethod(this, "Export", filename, action, format);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Export()
		{
			 Factory.ExecuteMethod(this, "Export");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Export(object filename)
		{
			 Factory.ExecuteMethod(this, "Export", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.SheetExportActionEnum Action = 1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Export(object filename, object action)
		{
			 Factory.ExecuteMethod(this, "Export", filename, action);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void LocateDataSource()
		{
			 Factory.ExecuteMethod(this, "LocateDataSource");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public object msDataSourceObject(string bstr)
		{
			return Factory.ExecuteVariantMethodGet(this, "msDataSourceObject", bstr);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="range1">NetOffice.OWC10Api._Range range1</param>
		/// <param name="range2">NetOffice.OWC10Api._Range range2</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range RectIntersect(NetOffice.OWC10Api._Range range1, NetOffice.OWC10Api._Range range2)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "RectIntersect", range1, range2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="range1">NetOffice.OWC10Api._Range range1</param>
		/// <param name="range2">NetOffice.OWC10Api._Range range2</param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range RectUnion(NetOffice.OWC10Api._Range range1, NetOffice.OWC10Api._Range range2)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api._Range>(this, "RectUnion", range1, range2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Repaint()
		{
			 Factory.ExecuteMethod(this, "Repaint");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ShowAbout()
		{
			 Factory.ExecuteMethod(this, "ShowAbout");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="menu">object menu</param>
		[SupportByVersion("OWC10", 1)]
		public void ShowContextMenu(Int32 x, Int32 y, object menu)
		{
			 Factory.ExecuteMethod(this, "ShowContextMenu", x, y, menu);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="topic">Int32 topic</param>
		[SupportByVersion("OWC10", 1)]
		public void ShowHelp(Int32 topic)
		{
			 Factory.ExecuteMethod(this, "ShowHelp", topic);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Undo()
		{
			 Factory.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void UpdatePropertyToolbox()
		{
			 Factory.ExecuteMethod(this, "UpdatePropertyToolbox");
		}

		#endregion

		#pragma warning restore
	}
}
