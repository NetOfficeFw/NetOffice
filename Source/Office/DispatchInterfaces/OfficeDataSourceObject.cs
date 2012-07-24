using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface OfficeDataSourceObject 
	/// SupportByVersion Office, 10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class OfficeDataSourceObject : COMObject
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
                    _type = typeof(OfficeDataSourceObject);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public string ConnectString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConnectString", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public string Table
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Table", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Table", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public string DataSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataSource", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataSource", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public object Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public Int32 RowCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public object Filters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Filters", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow MsoMoveRow</param>
		/// <param name="rowNbr">optional Int32 RowNbr = 1</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow, Int32 rowNbr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(msoMoveRow, rowNbr);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow MsoMoveRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(msoMoveRow);
			object returnItem = Invoker.MethodReturn(this, "Move", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		/// <param name="fNeverPrompt">optional Int32 fNeverPrompt = 1</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void Open(string bstrSrc, string bstrConnect, string bstrTable, Int32 fOpenExclusive, Int32 fNeverPrompt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSrc, bstrConnect, bstrTable, fOpenExclusive, fNeverPrompt);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void Open()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void Open(string bstrSrc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSrc);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void Open(string bstrSrc, string bstrConnect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSrc, bstrConnect);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void Open(string bstrSrc, string bstrConnect, string bstrTable)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSrc, bstrConnect, bstrTable);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void Open(string bstrSrc, string bstrConnect, string bstrTable, Int32 fOpenExclusive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrSrc, bstrConnect, bstrTable, fOpenExclusive);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="sortField1">string SortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		/// <param name="sortAscending3">optional bool SortAscending3 = true</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void SetSortOrder(string sortField1, bool sortAscending1, string sortField2, bool sortAscending2, string sortField3, bool sortAscending3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortField1, sortAscending1, sortField2, sortAscending2, sortField3, sortAscending3);
			Invoker.Method(this, "SetSortOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="sortField1">string SortField1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void SetSortOrder(string sortField1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortField1);
			Invoker.Method(this, "SetSortOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="sortField1">string SortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void SetSortOrder(string sortField1, bool sortAscending1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortField1, sortAscending1);
			Invoker.Method(this, "SetSortOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="sortField1">string SortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void SetSortOrder(string sortField1, bool sortAscending1, string sortField2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortField1, sortAscending1, sortField2);
			Invoker.Method(this, "SetSortOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="sortField1">string SortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void SetSortOrder(string sortField1, bool sortAscending1, string sortField2, bool sortAscending2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortField1, sortAscending1, sortField2, sortAscending2);
			Invoker.Method(this, "SetSortOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		/// <param name="sortField1">string SortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void SetSortOrder(string sortField1, bool sortAscending1, string sortField2, bool sortAscending2, string sortField3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sortField1, sortAscending1, sortField2, sortAscending2, sortField3);
			Invoker.Method(this, "SetSortOrder", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15)]
		public void ApplyFilter()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ApplyFilter", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}