using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface _Worksheet 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Worksheet : COMObject
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
                    _type = typeof(_Worksheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Worksheet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Worksheet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Worksheet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Worksheet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Worksheet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Worksheet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Worksheet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Worksheet(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821975.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196730.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192977.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837552.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836415.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196974.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836428.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198233.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834977.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834738.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837366.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197583.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834421.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197786.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821817.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837599.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool TransitionExpEval
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TransitionExpEval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TransitionExpEval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821903.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool AutoFilterMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFilterMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFilterMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841201.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194567.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Cells
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Cells", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197290.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range CircularReference
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "CircularReference", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197266.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Columns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Columns", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197853.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlConsolidationFunction ConsolidationFunction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlConsolidationFunction>(this, "ConsolidationFunction");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835571.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ConsolidationOptions
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ConsolidationOptions");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839655.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ConsolidationSources
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ConsolidationSources");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838608.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840106.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839855.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835599.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839763.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool FilterMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterMode");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197993.aspx </remarks>
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
		public string OnCalculate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnCalculate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnCalculate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnEntry
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnEntry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840285.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Outline Outline
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Outline>(this, "Outline", NetOffice.ExcelApi.Outline.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836512.aspx </remarks>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1, object cell2)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836512.aspx </remarks>
		/// <param name="cell1">object cell1</param>
		/// <param name="cell2">optional object cell2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public NetOffice.ExcelApi.Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836512.aspx </remarks>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.Range get_Range(object cell1)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836512.aspx </remarks>
		/// <param name="cell1">object cell1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public NetOffice.ExcelApi.Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821382.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Rows", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823064.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197479.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double StandardHeight
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "StandardHeight");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822174.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double StandardWidth
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "StandardWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StandardWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840554.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool TransitionFormEntry
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TransitionFormEntry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TransitionFormEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837858.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSheetType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSheetType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840732.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range UsedRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "UsedRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193761.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195715.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194075.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836199.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838771.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837757.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823161.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194187.aspx </remarks>
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
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196638.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836185.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197822.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197595.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Protection Protection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Protection>(this, "Protection", NetOffice.ExcelApi.Protection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195678.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObjects ListObjects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObjects>(this, "ListObjects", NetOffice.ExcelApi.ListObjects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840811.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195963.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196864.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838003.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837784.aspx </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Copy", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837784.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837784.aspx </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy(object before)
		{
			 Factory.ExecuteMethod(this, "Copy", before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837404.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834742.aspx </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Move", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834742.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move()
		{
			 Factory.ExecuteMethod(this, "Move");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834742.aspx </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move(object before)
		{
			 Factory.ExecuteMethod(this, "Move", before);
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840346.aspx </remarks>
		/// <param name="enableChanges">optional object enableChanges</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview(object enableChanges)
		{
			 Factory.ExecuteMethod(this, "PrintPreview", enableChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840346.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			 Factory.ExecuteMethod(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering, object allowUsingPivotTables)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTables });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect()
		{
			 Factory.ExecuteMethod(this, "Protect");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects)
		{
			 Factory.ExecuteMethod(this, "Protect", password, drawingObjects);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents)
		{
			 Factory.ExecuteMethod(this, "Protect", password, drawingObjects, contents);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			 Factory.ExecuteMethod(this, "Protect", password, drawingObjects, contents, scenarios);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840611.aspx </remarks>
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
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
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
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
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
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout, object local)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, fileFormat);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, fileFormat, password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword)
		{
			 Factory.ExecuteMethod(this, "SaveAs", filename, fileFormat, password, writeResPassword);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
		/// <param name="createBackup">optional object createBackup</param>
		/// <param name="addToMru">optional object addToMru</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195820.aspx </remarks>
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
		public void SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194988.aspx </remarks>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Select(object replace)
		{
			 Factory.ExecuteMethod(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194988.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841143.aspx </remarks>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Unprotect(object password)
		{
			 Factory.ExecuteMethod(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841143.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Unprotect()
		{
			 Factory.ExecuteMethod(this, "Unprotect");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821375.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void SetBackgroundPicture(string filename)
		{
			 Factory.ExecuteMethod(this, "SetBackgroundPicture", filename);
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834658.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Calculate()
		{
			 Factory.ExecuteMethod(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195149.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ChartObjects(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "ChartObjects", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195149.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194242.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194242.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling()
		{
			 Factory.ExecuteMethod(this, "CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194242.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194242.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194242.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196276.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ClearArrows()
		{
			 Factory.ExecuteMethod(this, "ClearArrows");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838386.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839053.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ResetAllPageBreaks()
		{
			 Factory.ExecuteMethod(this, "ResetAllPageBreaks");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197177.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OLEObjects(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "OLEObjects", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197177.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object OLEObjects()
		{
			return Factory.ExecuteVariantMethodGet(this, "OLEObjects");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821951.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		/// <param name="link">optional object link</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Paste(object destination, object link)
		{
			 Factory.ExecuteMethod(this, "Paste", destination, link);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821951.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821951.aspx </remarks>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Paste(object destination)
		{
			 Factory.ExecuteMethod(this, "Paste", destination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="noHTMLFormatting">optional object noHTMLFormatting</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object noHTMLFormatting)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex, iconLabel, noHTMLFormatting });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial()
		{
			 Factory.ExecuteMethod(this, "PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object format)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object format, object link)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object format, object link, object displayAsIcon)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object format, object link, object displayAsIcon, object iconFileName)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link, displayAsIcon, iconFileName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835858.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex });
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838199.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotTables(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotTables", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838199.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PivotTables()
		{
			return Factory.ExecuteVariantMethodGet(this, "PivotTables");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, sourceType, sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, sourceType, sourceData, tableDestination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, sourceType, sourceData, tableDestination, tableName);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
		/// <param name="sourceType">optional object sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="tableDestination">optional object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="rowGrand">optional object rowGrand</param>
		/// <param name="columnGrand">optional object columnGrand</param>
		/// <param name="saveData">optional object saveData</param>
		/// <param name="hasAutoFormat">optional object hasAutoFormat</param>
		/// <param name="autoPage">optional object autoPage</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839228.aspx </remarks>
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
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTableWizard", NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType, new object[]{ sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData });
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820786.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Scenarios(object index)
		{
			return Factory.ExecuteVariantMethodGet(this, "Scenarios", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820786.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Scenarios()
		{
			return Factory.ExecuteVariantMethodGet(this, "Scenarios");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197246.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ShowAllData()
		{
			 Factory.ExecuteMethod(this, "ShowAllData");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821077.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ShowDataForm()
		{
			 Factory.ExecuteMethod(this, "ShowDataForm");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823072.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ClearCircles()
		{
			 Factory.ExecuteMethod(this, "ClearCircles");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839372.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CircleInvalid()
		{
			 Factory.ExecuteMethod(this, "CircleInvalid");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821195.aspx </remarks>
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
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="spellLang">optional object spellLang</param>
		/// <param name="ignoreFinalYaa">optional object ignoreFinalYaa</param>
		/// <param name="spellScript">optional object spellScript</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang, object ignoreFinalYaa, object spellScript)
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa, spellScript });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _CheckSpelling()
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _CheckSpelling(object customDictionary)
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling", customDictionary, ignoreUppercase);
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
		public void _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
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
		public void _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
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
		public void _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang, object ignoreFinalYaa)
		{
			 Factory.ExecuteMethod(this, "_CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa });
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
		public void _Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			 Factory.ExecuteMethod(this, "_Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly });
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
		/// <param name="drawingObjects">optional object drawingObjects</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _Protect(object password, object drawingObjects)
		{
			 Factory.ExecuteMethod(this, "_Protect", password, drawingObjects);
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
		public void _Protect(object password, object drawingObjects, object contents)
		{
			 Factory.ExecuteMethod(this, "_Protect", password, drawingObjects, contents);
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
		public void _Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			 Factory.ExecuteMethod(this, "_Protect", password, drawingObjects, contents, scenarios);
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
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileFormat">optional object fileFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _SaveAs(string filename, object fileFormat)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename, fileFormat);
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
		public void _SaveAs(string filename, object fileFormat, object password)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename, fileFormat, password);
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
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", filename, fileFormat, password, writeResPassword);
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
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended });
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
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
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
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru });
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
		public void _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
		{
			 Factory.ExecuteMethod(this, "_SaveAs", new object[]{ filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage });
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
		public void _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _PasteSpecial()
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _PasteSpecial(object format)
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">optional object format</param>
		/// <param name="link">optional object link</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _PasteSpecial(object format, object link)
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial", format, link);
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
		public void _PasteSpecial(object format, object link, object displayAsIcon)
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial", format, link, displayAsIcon);
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
		public void _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName)
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial", format, link, displayAsIcon, iconFileName);
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
		public void _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex)
		{
			 Factory.ExecuteMethod(this, "_PasteSpecial", new object[]{ format, link, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839982.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespaces">optional object selectionNamespaces</param>
		/// <param name="map">optional object map</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range XmlDataQuery(string xPath, object selectionNamespaces, object map)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "XmlDataQuery", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, xPath, selectionNamespaces, map);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839982.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range XmlDataQuery(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "XmlDataQuery", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839982.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespaces">optional object selectionNamespaces</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range XmlDataQuery(string xPath, object selectionNamespaces)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "XmlDataQuery", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, xPath, selectionNamespaces);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837752.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespaces">optional object selectionNamespaces</param>
		/// <param name="map">optional object map</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range XmlMapQuery(string xPath, object selectionNamespaces, object map)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "XmlMapQuery", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, xPath, selectionNamespaces, map);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837752.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range XmlMapQuery(string xPath)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "XmlMapQuery", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, xPath);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837752.aspx </remarks>
		/// <param name="xPath">string xPath</param>
		/// <param name="selectionNamespaces">optional object selectionNamespaces</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range XmlMapQuery(string xPath, object selectionNamespaces)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Range>(this, "XmlMapQuery", NetOffice.ExcelApi.Range.LateBindingApiWrapperType, xPath, selectionNamespaces);
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840291.aspx </remarks>
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

		#endregion

		#pragma warning restore
	}
}
