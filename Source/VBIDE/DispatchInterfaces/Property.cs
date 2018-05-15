using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
    /// DispatchInterface Property
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface Property : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        object Value { get; set; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="index4">optional object index4</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_IndexedValue(object index1, object index2, object index3, object index4);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="index4">optional object index4</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_IndexedValue(object index1, object index2, object index3, object index4, object value);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="index4">optional object index4</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        object IndexedValue(object index1, object index2, object index3, object index4);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_IndexedValue(object index1);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="value">object value</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_IndexedValue(object index1, object value);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        object IndexedValue(object index1);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_IndexedValue(object index1, object index2);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_IndexedValue(object index1, object index2, object value);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        object IndexedValue(object index1, object index2);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_IndexedValue(object index1, object index2, object index3);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_IndexedValue(object index1, object index2, object index3, object value);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_IndexedValue
        /// </summary>
        /// <param name="index1">object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="index3">optional object index3</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_IndexedValue")]
        object IndexedValue(object index1, object index2, object index3);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int16 NumIndices { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VBIDEApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.VBIDEApi.Properties Parent { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        string Name { get; }

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
        NetOffice.VBIDEApi.Properties Collection { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3), ProxyResult]
        object Object { get; set; }

        #endregion
    }
}
