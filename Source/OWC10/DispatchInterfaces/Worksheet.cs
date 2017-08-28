using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface Worksheet 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Worksheet : COMObject
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
                    _type = typeof(Worksheet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Worksheet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Worksheet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheet(string progId) : base(progId)
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
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.AutoFilter AutoFilter
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.AutoFilter>(this, "AutoFilter", NetOffice.OWC10Api.AutoFilter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string CommandText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ConnectionString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionString", value);
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool FilterMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool IsDataBound
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDataBound");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsDataBound", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Next
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Next", NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Workbook Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Workbook>(this, "Parent", NetOffice.OWC10Api.Workbook.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Previous
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Previous", NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ProtectContents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectContents");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Protection Protection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Protection>(this, "Protection", NetOffice.OWC10Api.Protection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ProtectionMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProtectionMode");
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
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double StandardHeight
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "StandardHeight");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlSheetType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlSheetType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range UsedRange
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "UsedRange");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlSheetVisibility Visible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.XlSheetVisibility>(this, "Visible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Visible", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
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
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("OWC10", 1)]
		public void Copy(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Copy", before, after);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Copy(object before)
		{
			 Factory.ExecuteMethod(this, "Copy", before);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DumpStringTable()
		{
			 Factory.ExecuteMethod(this, "DumpStringTable");
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
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("OWC10", 1)]
		public void Move(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Move", before, after);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Move()
		{
			 Factory.ExecuteMethod(this, "Move");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Move(object before)
		{
			 Factory.ExecuteMethod(this, "Move", before);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="destination">optional object destination</param>
		/// <param name="link">optional object link</param>
		[SupportByVersion("OWC10", 1)]
		public void Paste(object destination, object link)
		{
			 Factory.ExecuteMethod(this, "Paste", destination, link);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Paste(object destination)
		{
			 Factory.ExecuteMethod(this, "Paste", destination);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		/// <param name="allowUsingPivotTableReports">optional object allowUsingPivotTableReports</param>
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering, object allowUsingPivotTableReports)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTableReports });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect()
		{
			 Factory.ExecuteMethod(this, "Protect");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password)
		{
			 Factory.ExecuteMethod(this, "Protect", password);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects)
		{
			 Factory.ExecuteMethod(this, "Protect", password, drawingObjects);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents)
		{
			 Factory.ExecuteMethod(this, "Protect", password, drawingObjects, contents);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			 Factory.ExecuteMethod(this, "Protect", password, drawingObjects, contents, scenarios);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		/// <param name="drawingObjects">optional object drawingObjects</param>
		/// <param name="contents">optional object contents</param>
		/// <param name="scenarios">optional object scenarios</param>
		/// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object allowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
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
		[SupportByVersion("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering)
		{
			 Factory.ExecuteMethod(this, "Protect", new object[]{ password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering });
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
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("OWC10", 1)]
		public void Select(object replace)
		{
			 Factory.ExecuteMethod(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ShowAllData()
		{
			 Factory.ExecuteMethod(this, "ShowAllData");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="password">optional object password</param>
		[SupportByVersion("OWC10", 1)]
		public void Unprotect(object password)
		{
			 Factory.ExecuteMethod(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Unprotect()
		{
			 Factory.ExecuteMethod(this, "Unprotect");
		}

		#endregion

		#pragma warning restore
	}
}
