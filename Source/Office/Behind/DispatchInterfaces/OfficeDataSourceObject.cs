using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface OfficeDataSourceObject 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864883.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class OfficeDataSourceObject : COMObject, NetOffice.OfficeApi.OfficeDataSourceObject
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.OfficeDataSourceObject);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public OfficeDataSourceObject() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861793.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string ConnectString
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectString");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectString", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861897.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string Table
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Table");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Table", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860869.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual string DataSource
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataSource");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataSource", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860229.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Columns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Columns");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861767.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 RowCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RowCount");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860598.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Filters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Filters");
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow, object rowNbr)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move", msoMoveRow, rowNbr);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864664.aspx </remarks>
        /// <param name="msoMoveRow">NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move(NetOffice.OfficeApi.Enums.MsoMoveRow msoMoveRow)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move", msoMoveRow);
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Open(object bstrSrc, object bstrConnect, object bstrTable, object fOpenExclusive, object fNeverPrompt)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Open", new object[] { bstrSrc, bstrConnect, bstrTable, fOpenExclusive, fNeverPrompt });
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Open()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Open");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Open(object bstrSrc)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Open", bstrSrc);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        /// <param name="bstrConnect">optional string bstrConnect = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Open(object bstrSrc, object bstrConnect)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Open", bstrSrc, bstrConnect);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865289.aspx </remarks>
        /// <param name="bstrSrc">optional string bstrSrc = </param>
        /// <param name="bstrConnect">optional string bstrConnect = </param>
        /// <param name="bstrTable">optional string bstrTable = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Open(object bstrSrc, object bstrConnect, object bstrTable)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Open", bstrSrc, bstrConnect, bstrTable);
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void Open(object bstrSrc, object bstrConnect, object bstrTable, object fOpenExclusive)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Open", bstrSrc, bstrConnect, bstrTable, fOpenExclusive);
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3, object sortAscending3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", new object[] { sortField1, sortAscending1, sortField2, sortAscending2, sortField3, sortAscending3 });
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetSortOrder(string sortField1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetSortOrder(string sortField1, object sortAscending1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861392.aspx </remarks>
        /// <param name="sortField1">string sortField1</param>
        /// <param name="sortAscending1">optional bool SortAscending1 = true</param>
        /// <param name="sortField2">optional string SortField2 = </param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetSortOrder(string sortField1, object sortAscending1, object sortField2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2);
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", sortField1, sortAscending1, sortField2, sortAscending2);
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
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void SetSortOrder(string sortField1, object sortAscending1, object sortField2, object sortAscending2, object sortField3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSortOrder", new object[] { sortField1, sortAscending1, sortField2, sortAscending2, sortField3 });
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863341.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void ApplyFilter()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter");
        }

        #endregion

        #pragma warning restore
    }
}
