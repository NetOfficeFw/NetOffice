using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface OfficeDataSourceObject 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864883.aspx </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OfficeDataSourceObject : COMObject
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
                    _type = typeof(OfficeDataSourceObject);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public OfficeDataSourceObject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OfficeDataSourceObject(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OfficeDataSourceObject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861793.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string ConnectString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectString", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861897.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string Table
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Table");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Table", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860869.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string DataSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860229.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16), ProxyResult]
		public object Columns
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Columns");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861767.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 RowCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RowCount");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860598.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16), ProxyResult]
		public object Filters
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Filters");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864664.aspx </remarks>
		/// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow</param>
		/// <param name="rowNbr">optional Int32 RowNbr = 1</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow, object rowNbr)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", msoMoveRow, rowNbr);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864664.aspx </remarks>
		/// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", msoMoveRow);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		/// <param name="fNeverPrompt">optional Int32 fNeverPrompt = 1</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Open(object bstrSrc, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt)
		{
			 Factory.ExecuteMethod(this, "Open", new object[]{ bstrSrc, bstrConnect, bstrTable, fOpenExclusive, fNeverPrompt });
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Open(object bstrSrc)
		{
			 Factory.ExecuteMethod(this, "Open", bstrSrc);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Open(object bstrSrc, object bstrConnect)
		{
			 Factory.ExecuteMethod(this, "Open", bstrSrc, bstrConnect);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Open(object bstrSrc, object bstrConnect, object bstrTable)
		{
			 Factory.ExecuteMethod(this, "Open", bstrSrc, bstrConnect, bstrTable);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
		/// <param name="bstrSrc">optional string bstrSrc = </param>
		/// <param name="bstrConnect">optional string bstrConnect = </param>
		/// <param name="bstrTable">optional string bstrTable = </param>
		/// <param name="fOpenExclusive">optional Int32 fOpenExclusive = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Open(object bstrSrc, object bstrConnect, object bstrTable, object fOpenExclusive)
		{
			 Factory.ExecuteMethod(this, "Open", bstrSrc, bstrConnect, bstrTable, fOpenExclusive);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
		/// <param name="sortField1">string sortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		/// <param name="sortAscending3">optional bool SortAscending3 = true</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3, object sortAscending3)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", new object[]{ sortField1, sortAscending1, sortField2, sortAscending2, sortField3, sortAscending3 });
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
		/// <param name="sortField1">string sortField1</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetSortOrder(string sortField1)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
		/// <param name="sortField1">string sortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetSortOrder(string sortField1, object sortAscending1)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
		/// <param name="sortField1">string sortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetSortOrder(string sortField1, object sortAscending1, object sortField2)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
		/// <param name="sortField1">string sortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2, sortAscending2);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
		/// <param name="sortField1">string sortField1</param>
		/// <param name="sortAscending1">optional bool SortAscending1 = true</param>
		/// <param name="sortField2">optional string SortField2 = </param>
		/// <param name="sortAscending2">optional bool SortAscending2 = true</param>
		/// <param name="sortField3">optional string SortField3 = </param>
		[CustomMethod]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3)
		{
			 Factory.ExecuteMethod(this, "SetSortOrder", new object[]{ sortField1, sortAscending1, sortField2, sortAscending2, sortField3 });
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863341.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void ApplyFilter()
		{
			 Factory.ExecuteMethod(this, "ApplyFilter");
		}

		#endregion

		#pragma warning restore
	}
}
