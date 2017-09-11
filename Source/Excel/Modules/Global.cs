using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi.GlobalHelperModules
{
    ///<summary>
    /// Module GlobalModule
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    ///</summary>
    [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsModule), ModuleBaseType(typeof(ExcelApi.Application))]
	public static class GlobalModule
	{
		#region Fields

		private static ICOMObject _instance;

        #endregion

        #region Internal Properties

        internal static ICOMObject Instance
        {
            get
            {
                return _instance;
            }
            set
            {
                if ((null == value) || (null == _instance))
                    _instance = value;
            }
        }

        internal static Core Factory
		{
			get
			{
				if(null != _instance)
					 return _instance.Factory;
			else
				return Core.Default;
			}
		}

		internal static Invoker Invoker
		{
			get
			{
				if(null != _instance)
					 return _instance.Invoker;
			else
				return Invoker.Default;
			}
		}

		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Application Application
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(_instance, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(_instance, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Application Parent
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(_instance, "Parent", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
            }
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range ActiveCell
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(_instance, "ActiveCell", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Chart ActiveChart
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Chart>(_instance, "ActiveChart", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.DialogSheet ActiveDialog
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DialogSheet>(_instance, "ActiveDialog", NetOffice.ExcelApi.DialogSheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.MenuBar ActiveMenuBar
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.MenuBar>(_instance, "ActiveMenuBar", NetOffice.ExcelApi.MenuBar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static string ActivePrinter
		{
			get
			{
                return Factory.ExecuteStringPropertyGet(_instance, "ActivePrinter");
			}
			set
			{
                Factory.ExecuteValuePropertySet(_instance, "ActivePrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public static object ActiveSheet
		{
			get
			{
                return Factory.ExecuteReferencePropertyGet(ThisWorkbook, "ActiveSheet");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Window ActiveWindow
		{
			get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Window>(_instance, "ActiveWindow", NetOffice.ExcelApi.Window.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Workbook ActiveWorkbook
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Workbook>(_instance, "ActiveWorkbook", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.AddIns AddIns
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AddIns>(_instance, "AddIns", NetOffice.ExcelApi.AddIns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.Assistant Assistant
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(_instance, "Assistant", NetOffice.OfficeApi.Assistant.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Cells
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(_instance, "Cells", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Sheets Charts
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(_instance, "Charts", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Range Columns
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(_instance, "Columns", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(_instance, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static Int32 DDEAppReturnCode
		{
			get
			{
                return Factory.ExecuteInt32PropertyGet(_instance, "DDEAppReturnCode");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Sheets DialogSheets
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(_instance, "DialogSheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.MenuBars MenuBars
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.MenuBars>(_instance, "MenuBars", NetOffice.ExcelApi.MenuBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Modules Modules
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Modules>(_instance, "Modules", NetOffice.ExcelApi.Modules.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Names Names
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Names>(_instance, "Names", NetOffice.ExcelApi.Names.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Range get_Range(object cell1, object cell2)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(_instance, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public static NetOffice.ExcelApi.Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Range get_Range(object cell1)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(_instance, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public static NetOffice.ExcelApi.Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Range Rows
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(_instance, "Rows", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public static object Selection
		{
			get
			{
                return Factory.ExecuteReferencePropertyGet(_instance, "Selection");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Sheets Sheets
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(_instance, "Sheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Menu get_ShortcutMenus(Int32 index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Menu>(_instance, "ShortcutMenus", NetOffice.ExcelApi.Menu.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ShortcutMenus
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_ShortcutMenus")]
		public static NetOffice.ExcelApi.Menu ShortcutMenus(Int32 index)
		{
			return get_ShortcutMenus(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Workbook ThisWorkbook
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Workbook>(_instance, "ThisWorkbook", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public static NetOffice.ExcelApi.Toolbars Toolbars
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Toolbars>(_instance, "Toolbars", NetOffice.ExcelApi.Toolbars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Windows Windows
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Windows>(_instance, "Windows", NetOffice.ExcelApi.Windows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Workbooks Workbooks
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Workbooks>(_instance, "Workbooks", NetOffice.ExcelApi.Workbooks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.WorksheetFunction WorksheetFunction
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorksheetFunction>(_instance, "WorksheetFunction", NetOffice.ExcelApi.WorksheetFunction.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Sheets Worksheets
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(_instance, "Worksheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Sheets Excel4IntlMacroSheets
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(_instance, "Excel4IntlMacroSheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Sheets Excel4MacroSheets
		{
			get
			{
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(_instance, "Excel4MacroSheets", NetOffice.ExcelApi.Sheets.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static void Calculate()
		{
            Factory.ExecuteMethod(_instance, "Calculate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="channel">Int32 channel</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static void DDEExecute(Int32 channel, string _string)
		{
            Factory.ExecuteMethod(_instance, "DDEExecute", channel, _string);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="app">string app</param>
		/// <param name="topic">string topic</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static Int32 DDEInitiate(string app, string topic)
		{
            return Factory.ExecuteInt32MethodGet(_instance, "DDEInitiate", app, topic);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="channel">Int32 channel</param>
		/// <param name="item">object item</param>
		/// <param name="data">object data</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static void DDEPoke(Int32 channel, object item, object data)
		{
            Factory.ExecuteMethod(_instance, "DDEPoke", channel, item, data);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="channel">Int32 channel</param>
		/// <param name="item">string item</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object DDERequest(Int32 channel, string item)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Selection", channel, item);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="channel">Int32 channel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static void DDETerminate(Int32 channel)
		{
            Factory.ExecuteMethod(_instance, "DDETerminate", channel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">object name</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Evaluate(object name)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Evaluate", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="name">object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Evaluate(object name)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Evaluate", name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="_string">string string</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object ExecuteExcel4Macro(string _string)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "ExecuteExcel4Macro", _string);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		/// <param name="arg30">optional object arg30</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5, arg6);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Intersect(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Intersect", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		/// <param name="arg30">optional object arg30</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run()
        {
            return Factory.ExecuteVariantMethodGet(_instance, "Run");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2, arg3);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2, arg3, arg4, arg5);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2, arg3, arg4, arg5, arg6);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        /// <param name="arg29">optional object arg29</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object Run(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        /// <param name="arg29">optional object arg29</param>
        /// <param name="arg30">optional object arg30</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "Run", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2()
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1, arg2, arg3);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1, arg2, arg3, arg4);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1, arg2, arg3, arg4, arg5);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1, arg2, arg3, arg4, arg5, arg6);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macro">optional object macro</param>
        /// <param name="arg1">optional object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="macro">optional object macro</param>
		/// <param name="arg1">optional object arg1</param>
		/// <param name="arg2">optional object arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static object _Run2(object macro, object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
            return Factory.ExecuteVariantMethodGet(_instance, "_Run2", new object[] { macro, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="keys">object keys</param>
        /// <param name="wait">optional object wait</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static void SendKeys(object keys, object wait)
		{
            Factory.ExecuteMethod(_instance, "SendKeys", keys, wait);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="keys">object keys</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static void SendKeys(object keys)
		{
            Factory.ExecuteMethod(_instance, "SendKeys", keys);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		/// <param name="arg30">optional object arg30</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29, object arg30)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29, arg30 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5, arg6);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5, arg6, arg7);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8);
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16});
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        /// <param name="arg10">optional object arg10</param>
        /// <param name="arg11">optional object arg11</param>
        /// <param name="arg12">optional object arg12</param>
        /// <param name="arg13">optional object arg13</param>
        /// <param name="arg14">optional object arg14</param>
        /// <param name="arg15">optional object arg15</param>
        /// <param name="arg16">optional object arg16</param>
        /// <param name="arg17">optional object arg17</param>
        /// <param name="arg18">optional object arg18</param>
        /// <param name="arg19">optional object arg19</param>
        /// <param name="arg20">optional object arg20</param>
        /// <param name="arg21">optional object arg21</param>
        /// <param name="arg22">optional object arg22</param>
        /// <param name="arg23">optional object arg23</param>
        /// <param name="arg24">optional object arg24</param>
        /// <param name="arg25">optional object arg25</param>
        /// <param name="arg26">optional object arg26</param>
        /// <param name="arg27">optional object arg27</param>
        /// <param name="arg28">optional object arg28</param>
        [CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28 });
        }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
		/// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
		/// <param name="arg3">optional object arg3</param>
		/// <param name="arg4">optional object arg4</param>
		/// <param name="arg5">optional object arg5</param>
		/// <param name="arg6">optional object arg6</param>
		/// <param name="arg7">optional object arg7</param>
		/// <param name="arg8">optional object arg8</param>
		/// <param name="arg9">optional object arg9</param>
		/// <param name="arg10">optional object arg10</param>
		/// <param name="arg11">optional object arg11</param>
		/// <param name="arg12">optional object arg12</param>
		/// <param name="arg13">optional object arg13</param>
		/// <param name="arg14">optional object arg14</param>
		/// <param name="arg15">optional object arg15</param>
		/// <param name="arg16">optional object arg16</param>
		/// <param name="arg17">optional object arg17</param>
		/// <param name="arg18">optional object arg18</param>
		/// <param name="arg19">optional object arg19</param>
		/// <param name="arg20">optional object arg20</param>
		/// <param name="arg21">optional object arg21</param>
		/// <param name="arg22">optional object arg22</param>
		/// <param name="arg23">optional object arg23</param>
		/// <param name="arg24">optional object arg24</param>
		/// <param name="arg25">optional object arg25</param>
		/// <param name="arg26">optional object arg26</param>
		/// <param name="arg27">optional object arg27</param>
		/// <param name="arg28">optional object arg28</param>
		/// <param name="arg29">optional object arg29</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public static NetOffice.ExcelApi.Range Union(NetOffice.ExcelApi.Range arg1, NetOffice.ExcelApi.Range arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9, object arg10, object arg11, object arg12, object arg13, object arg14, object arg15, object arg16, object arg17, object arg18, object arg19, object arg20, object arg21, object arg22, object arg23, object arg24, object arg25, object arg26, object arg27, object arg28, object arg29)
		{
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(_instance, "Union", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16, arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24, arg25, arg26, arg27, arg28, arg29 });
        }

		#endregion
	}
}
