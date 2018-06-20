using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface Scripts 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("000C0340-0000-0000-C000-000000000046")]
    public interface Scripts : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.Script>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.Script this[object index] { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        /// <param name="id">optional string Id = </param>
        /// <param name="extended">optional string Extended = </param>
        /// <param name="scriptText">optional string ScriptText = </param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended, object scriptText);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add();

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add(object anchor);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add(object anchor, object location);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add(object anchor, object location, object language);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        /// <param name="id">optional string Id = </param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        /// <param name="id">optional string Id = </param>
        /// <param name="extended">optional string Extended = </param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Delete();

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.Script>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.Script> GetEnumerator();

        #endregion
    }
}
