using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IModel 
    /// SupportByVersion Excel, 15, 16
    /// </summary>
    [SupportByVersion("Excel", 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IModel : COMObject, NetOffice.ExcelApi.IModel
    {
        #pragma warning disable

        #region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IModel);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <param name="factory">current used factory core</param>
        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="proxyShare">proxy share instead if com proxy</param>
        public IModel(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
        {
        }

        ///<param name="factory">current used factory core</param>
        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        public IModel(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
        {

        }

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IModel(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
        {
        }

        ///<param name="factory">current used factory core</param>
        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IModel(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
        {

        }

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IModel(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
        {
        }

        ///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IModel(ICOMObject replacedObject) : base(replacedObject)
        {
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IModel() : base()
        {
        }

        /// <param name="progId">registered progID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IModel(string progId) : base(progId)
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.ModelTables ModelTables
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelTables>(this, "ModelTables", typeof(NetOffice.ExcelApi.ModelTables));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.ModelRelationships ModelRelationships
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelRelationships>(this, "ModelRelationships", typeof(NetOffice.ExcelApi.ModelRelationships));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.WorkbookConnection DataModelConnection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookConnection>(this, "DataModelConnection", typeof(NetOffice.ExcelApi.WorkbookConnection));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public Int32 Refresh()
        {
            return Factory.ExecuteInt32MethodGet(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="connectionToDataSource">NetOffice.ExcelApi.WorkbookConnection connectionToDataSource</param>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.WorkbookConnection AddConnection(NetOffice.ExcelApi.WorkbookConnection connectionToDataSource)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "AddConnection", typeof(NetOffice.ExcelApi.WorkbookConnection), connectionToDataSource);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="modelTable">object modelTable</param>
        [SupportByVersion("Excel", 15, 16)]
        public NetOffice.ExcelApi.WorkbookConnection CreateModelWorkbookConnection(object modelTable)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.WorkbookConnection>(this, "CreateModelWorkbookConnection", typeof(NetOffice.ExcelApi.WorkbookConnection), modelTable);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        public Int32 Initialize()
        {
            return Factory.ExecuteInt32MethodGet(this, "Initialize");
        }

        #endregion

        #pragma warning restore
    }
}

