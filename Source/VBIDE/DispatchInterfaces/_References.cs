using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface _References
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "VBIDE", 12, 14, 5.3), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0002E17A-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VBIDEApi.References))]
    public interface _References : ICOMObject, IEnumerableProvider<NetOffice.VBIDEApi.Reference>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBProject Parent { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBE VBE { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.VBIDEApi.Reference this[object index] { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="guid">string guid</param>
        /// <param name="major">Int32 major</param>
        /// <param name="minor">Int32 minor</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Reference AddFromGuid(string guid, Int32 major, Int32 minor);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Reference AddFromFile(string fileName);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="reference">NetOffice.VBIDEApi.Reference reference</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void Remove(NetOffice.VBIDEApi.Reference reference);

        #endregion

        #region IEnumerable<NetOffice.VBIDEApi.Reference>

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        new IEnumerator<NetOffice.VBIDEApi.Reference> GetEnumerator();

        #endregion
    }
}
