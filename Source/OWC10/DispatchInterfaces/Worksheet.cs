using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface Worksheet 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Worksheet : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OWC10Api.ISpreadsheet newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api.ISpreadsheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.AutoFilter AutoFilter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoFilter", paramsArray);
				NetOffice.OWC10Api.AutoFilter newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.AutoFilter.LateBindingApiWrapperType) as NetOffice.OWC10Api.AutoFilter;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool AutoFilterMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoFilterMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoFilterMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string CommandText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CommandText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string ConnectionString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConnectionString", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string DataMember
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataMember", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataMember", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool EnableAutoFilter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EnableAutoFilter", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "EnableAutoFilter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool FilterMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FilterMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool IsDataBound
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsDataBound", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IsDataBound", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Names Names
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Names", paramsArray);
				NetOffice.OWC10Api.Names newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Names.LateBindingApiWrapperType) as NetOffice.OWC10Api.Names;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Next
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Next", paramsArray);
				NetOffice.OWC10Api.Worksheet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType) as NetOffice.OWC10Api.Worksheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Workbook Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.OWC10Api.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Workbook.LateBindingApiWrapperType) as NetOffice.OWC10Api.Workbook;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Worksheet Previous
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Previous", paramsArray);
				NetOffice.OWC10Api.Worksheet newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Worksheet.LateBindingApiWrapperType) as NetOffice.OWC10Api.Worksheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool ProtectContents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectContents", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Protection Protection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Protection", paramsArray);
				NetOffice.OWC10Api.Protection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Protection.LateBindingApiWrapperType) as NetOffice.OWC10Api.Protection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool ProtectionMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProtectionMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1, object cell2)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1, cell2);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		/// <param name="cell2">optional object Cell2</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Range(object cell1, object cell2)
		{
			return get_Range(cell1, cell2);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api._Range get_Range(object cell1)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cell1);
			object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
			NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Range
		/// </summary>
		/// <param name="cell1">object Cell1</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Range(object cell1)
		{
			return get_Range(cell1);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Double StandardHeight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StandardHeight", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Double StandardWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StandardWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StandardWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlSheetType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.XlSheetType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api._Range UsedRange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UsedRange", paramsArray);
				NetOffice.OWC10Api._Range newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OWC10Api._Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.XlSheetVisibility Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.XlSheetVisibility)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Calculate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Calculate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="before">optional object Before</param>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Copy(object before, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before, after);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Copy(object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void DumpStringTable()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DumpStringTable", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="expression">object Expression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public object _Evaluate(object expression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expression);
			object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="expression">object Expression</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Evaluate(object expression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(expression);
			object returnItem = Invoker.MethodReturn(this, "Evaluate", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="before">optional object Before</param>
		/// <param name="after">optional object After</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Move(object before, object after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before, after);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Move()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Move(object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(before);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		/// <param name="link">optional object Link</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Paste(object destination, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination, link);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="destination">optional object Destination</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Paste(object destination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destination);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object AllowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object AllowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object AllowDeletingRows</param>
		/// <param name="allowSorting">optional object AllowSorting</param>
		/// <param name="allowFiltering">optional object AllowFiltering</param>
		/// <param name="allowUsingPivotTableReports">optional object AllowUsingPivotTableReports</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering, object allowUsingPivotTableReports)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTableReports);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object AllowInsertingHyperlinks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object AllowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object AllowDeletingColumns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object AllowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object AllowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object AllowDeletingRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object AllowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object AllowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object AllowDeletingRows</param>
		/// <param name="allowSorting">optional object AllowSorting</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		/// <param name="drawingObjects">optional object DrawingObjects</param>
		/// <param name="contents">optional object Contents</param>
		/// <param name="scenarios">optional object Scenarios</param>
		/// <param name="userInterfaceOnly">optional object UserInterfaceOnly</param>
		/// <param name="allowFormattingCells">optional object AllowFormattingCells</param>
		/// <param name="allowFormattingColumns">optional object AllowFormattingColumns</param>
		/// <param name="allowFormattingRows">optional object AllowFormattingRows</param>
		/// <param name="allowInsertingColumns">optional object AllowInsertingColumns</param>
		/// <param name="allowInsertingRows">optional object AllowInsertingRows</param>
		/// <param name="allowInsertingHyperlinks">optional object AllowInsertingHyperlinks</param>
		/// <param name="allowDeletingColumns">optional object AllowDeletingColumns</param>
		/// <param name="allowDeletingRows">optional object AllowDeletingRows</param>
		/// <param name="allowSorting">optional object AllowSorting</param>
		/// <param name="allowFiltering">optional object AllowFiltering</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering);
			Invoker.Method(this, "Protect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Refresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="replace">optional object Replace</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Select(object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(replace);
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void ShowAllData()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ShowAllData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Unprotect(object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(password);
			Invoker.Method(this, "Unprotect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Unprotect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Unprotect", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}