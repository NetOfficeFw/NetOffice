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
	/// Interface IProtectedViewWindows 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000244CC-0001-0000-C000-000000000046")]
	public interface IProtectedViewWindows : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.ProtectedViewWindow>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.ProtectedViewWindow this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="repairMode">optional object repairMode</param>
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ProtectedViewWindow Open(string filename, object password, object addToMru, object repairMode);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ProtectedViewWindow Open(string filename);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ProtectedViewWindow Open(string filename, object password);

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="password">optional object password</param>
		/// <param name="addToMru">optional object addToMru</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		NetOffice.ExcelApi.ProtectedViewWindow Open(string filename, object password, object addToMru);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.ProtectedViewWindow>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.ProtectedViewWindow> GetEnumerator();

        #endregion
    }
}
