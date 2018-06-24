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
	/// DispatchInterface ToolbarButtons 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("0002085F-0000-0000-C000-000000000046")]
	public interface ToolbarButtons : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.ToolbarButton>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.ToolbarButton this[Int32 index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="pushed">optional object pushed</param>
		/// <param name="enabled">optional object enabled</param>
		/// <param name="statusBar">optional object statusBar</param>
		/// <param name="helpFile">optional object helpFile</param>
		/// <param name="helpContextID">optional object helpContextID</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before, object onAction, object pushed, object enabled, object statusBar, object helpFile, object helpContextID);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		/// <param name="onAction">optional object onAction</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before, object onAction);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="pushed">optional object pushed</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before, object onAction, object pushed);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="pushed">optional object pushed</param>
		/// <param name="enabled">optional object enabled</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before, object onAction, object pushed, object enabled);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="pushed">optional object pushed</param>
		/// <param name="enabled">optional object enabled</param>
		/// <param name="statusBar">optional object statusBar</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before, object onAction, object pushed, object enabled, object statusBar);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="button">optional object button</param>
		/// <param name="before">optional object before</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="pushed">optional object pushed</param>
		/// <param name="enabled">optional object enabled</param>
		/// <param name="statusBar">optional object statusBar</param>
		/// <param name="helpFile">optional object helpFile</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.ToolbarButton Add(object button, object before, object onAction, object pushed, object enabled, object statusBar, object helpFile);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.ToolbarButton>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.ToolbarButton> GetEnumerator();

        #endregion
    }
}
