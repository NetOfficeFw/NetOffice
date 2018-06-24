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
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196294.aspx </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("00024486-0000-0000-C000-000000000046")]
	public interface Connections : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.WorkbookConnection>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838431.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840518.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834962.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839189.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.WorkbookConnection this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195309.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195309.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="createModelConnection">optional object createModelConnection</param>
		/// <param name="importRelationships">optional object importRelationships</param>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename, object createModelConnection, object importRelationships);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195309.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="createModelConnection">optional object createModelConnection</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.WorkbookConnection AddFromFile(string filename, object createModelConnection);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="description">string description</param>
		/// <param name="connectionString">object connectionString</param>
		/// <param name="commandText">object commandText</param>
		/// <param name="lCmdtype">optional object lCmdtype</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype);

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
		NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype, object createModelConnection, object importRelationships);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="description">string description</param>
		/// <param name="connectionString">object connectionString</param>
		/// <param name="commandText">object commandText</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText);

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
		NetOffice.ExcelApi.WorkbookConnection Add(string name, string description, object connectionString, object commandText, object lCmdtype, object createModelConnection);

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.WorkbookConnection _AddFromFile(string filename);

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
		NetOffice.ExcelApi.WorkbookConnection _Add(string name, string description, object connectionString, object commandText, object lCmdtype);

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
		NetOffice.ExcelApi.WorkbookConnection _Add(string name, string description, object connectionString, object commandText);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.WorkbookConnection>

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.WorkbookConnection> GetEnumerator();

        #endregion
    }
}
