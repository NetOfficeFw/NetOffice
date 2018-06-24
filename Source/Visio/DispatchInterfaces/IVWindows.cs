using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVWindows 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Visio", 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000D0711-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Windows))]
    public interface IVWindows : ICOMObject, IEnumerableProvider<NetOffice.VisioApi.IVWindow>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.VisioApi.IVWindow this[Int16 index] { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Count { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 PersistsEvents { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVWindow get_ItemFromID(Int32 nID);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemFromID")]
		NetOffice.VisioApi.IVWindow ItemFromID(Int32 nID);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="captionOrIndex">object captionOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVWindow get_ItemEx(object captionOrIndex);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemEx
		/// </summary>
		/// <param name="captionOrIndex">object captionOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemEx")]
		NetOffice.VisioApi.IVWindow ItemEx(object captionOrIndex);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void VoidArrange();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nArrangeFlags">optional object nArrangeFlags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Arrange(object nArrangeFlags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Arrange();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		/// <param name="bstrMergeClass">optional object bstrMergeClass</param>
		/// <param name="nMergePosition">optional object nMergePosition</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass, object nMergePosition);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		/// <param name="bstrMergeClass">optional object bstrMergeClass</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass);

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVWindow>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.VisioApi.IVWindow> GetEnumerator();

        #endregion
    }
}
