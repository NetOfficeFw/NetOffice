using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IConnections 
    /// SupportByVersion Excel, 12,14,15,16
    /// </summary>
    public class IConnections : COMObject, NetOffice.ExcelApi.IConnections
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
                    _contractType = typeof(NetOffice.ExcelApi.IConnections);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type        /// </summary>
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
                    _type = typeof(IConnections);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IConnections() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.ExcelApi.WorkbookConnection this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_Default", typeof(NetOffice.ExcelApi.WorkbookConnection), index);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddFromFile", typeof(NetOffice.ExcelApi.WorkbookConnection), filename);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="createModelConnection">optional object createModelConnection</param>
        /// <param name="importRelationships">optional object importRelationships</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename, object createModelConnection, object importRelationships)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddFromFile", typeof(NetOffice.ExcelApi.WorkbookConnection), filename, createModelConnection, importRelationships);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="createModelConnection">optional object createModelConnection</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename, object createModelConnection)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddFromFile", typeof(NetOffice.ExcelApi.WorkbookConnection), filename, createModelConnection);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="description">string description</param>
        /// <param name="connectionString">object connectionString</param>
        /// <param name="commandText">object commandText</param>
        /// <param name="lCmdtype">optional object lCmdtype</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", typeof(NetOffice.ExcelApi.WorkbookConnection), new object[] { name, description, connectionString, commandText, lCmdtype });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="description">string description</param>
        /// <param name="connectionString">object connectionString</param>
        /// <param name="commandText">object commandText</param>
        /// <param name="lCmdtype">optional object lCmdtype</param>
        /// <param name="createModelConnection">optional object createModelConnection</param>
        /// <param name="importRelationships">optional object importRelationships</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype, object createModelConnection, object importRelationships)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", typeof(NetOffice.ExcelApi.WorkbookConnection), new object[] { name, description, connectionString, commandText, lCmdtype, createModelConnection, importRelationships });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="description">string description</param>
        /// <param name="connectionString">object connectionString</param>
        /// <param name="commandText">object commandText</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", typeof(NetOffice.ExcelApi.WorkbookConnection), name, description, connectionString, commandText);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="description">string description</param>
        /// <param name="connectionString">object connectionString</param>
        /// <param name="commandText">object commandText</param>
        /// <param name="lCmdtype">optional object lCmdtype</param>
        /// <param name="createModelConnection">optional object createModelConnection</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype, object createModelConnection)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", typeof(NetOffice.ExcelApi.WorkbookConnection), new object[] { name, description, connectionString, commandText, lCmdtype, createModelConnection });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="filename">string filename</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection _AddFromFile(string filename)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_AddFromFile", typeof(NetOffice.ExcelApi.WorkbookConnection), filename);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="description">string description</param>
        /// <param name="connectionString">object connectionString</param>
        /// <param name="commandText">object commandText</param>
        /// <param name="lCmdtype">optional object lCmdtype</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection _Add(string name, string description, object connectionString, object commandText, object lCmdtype)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_Add", typeof(NetOffice.ExcelApi.WorkbookConnection), new object[] { name, description, connectionString, commandText, lCmdtype });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="description">string description</param>
        /// <param name="connectionString">object connectionString</param>
        /// <param name="commandText">object commandText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.WorkbookConnection _Add(string name, string description, object connectionString, object commandText)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_Add", typeof(NetOffice.ExcelApi.WorkbookConnection), name, description, connectionString, commandText);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.WorkbookConnection>

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.WorkbookConnection> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.WorkbookConnection item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}

