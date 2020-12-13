using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Connections 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections"/> </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class Connections : COMObject, IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>
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
                    _type = typeof(Connections);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Connections(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Connections(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connections(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connections(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connections(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connections(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connections() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Connections(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.Application"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.Creator"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.Parent"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.Count"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.WorkbookConnection this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_Default", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.AddFromFile"/> </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddFromFile", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.AddFromFile"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="createModelConnection">optional object createModelConnection</param>
		/// <param name="importRelationships">optional object importRelationships</param>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename, object createModelConnection, object importRelationships)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddFromFile", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, filename, createModelConnection, importRelationships);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Connections.AddFromFile"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="createModelConnection">optional object createModelConnection</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename, object createModelConnection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddFromFile", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, filename, createModelConnection);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="description">string description</param>
		/// <param name="connectionString">object connectionString</param>
		/// <param name="commandText">object commandText</param>
		/// <param name="lCmdtype">optional object lCmdtype</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, new object[]{ name, description, connectionString, commandText, lCmdtype });
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
		public NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype, object createModelConnection, object importRelationships)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, new object[]{ name, description, connectionString, commandText, lCmdtype, createModelConnection, importRelationships });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="description">string description</param>
		/// <param name="connectionString">object connectionString</param>
		/// <param name="commandText">object commandText</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, name, description, connectionString, commandText);
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
		public NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype, object createModelConnection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "Add", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, new object[]{ name, description, connectionString, commandText, lCmdtype, createModelConnection });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.WorkbookConnection _AddFromFile(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_AddFromFile", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, filename);
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
		public NetOffice.ExcelApi.WorkbookConnection _Add(string name, string description, object connectionString, object commandText, object lCmdtype)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_Add", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, new object[]{ name, description, connectionString, commandText, lCmdtype });
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
		public NetOffice.ExcelApi.WorkbookConnection _Add(string name, string description, object connectionString, object commandText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "_Add", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType, name, description, connectionString, commandText);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.WorkbookConnection> Member

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.WorkbookConnection> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.WorkbookConnection item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}
