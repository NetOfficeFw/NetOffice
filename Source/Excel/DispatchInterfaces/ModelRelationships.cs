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
	/// DispatchInterface ModelRelationships 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231609.aspx </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 15, 16), HasIndexProperty(IndexInvoke.Property, "_Default")]
	[TypeId("000244DA-0000-0000-C000-000000000046")]
	public interface ModelRelationships : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.ModelRelationship>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230687.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231382.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232120.aspx </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230559.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.ModelRelationship this[object index] { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230789.aspx </remarks>
		/// <param name="foreignKeyColumn">NetOffice.ExcelApi.ModelTableColumn foreignKeyColumn</param>
		/// <param name="primaryKeyColumn">NetOffice.ExcelApi.ModelTableColumn primaryKeyColumn</param>
		[SupportByVersion("Excel", 15, 16)]
		NetOffice.ExcelApi.ModelRelationship Add(NetOffice.ExcelApi.ModelTableColumn foreignKeyColumn, NetOffice.ExcelApi.ModelTableColumn primaryKeyColumn);

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.ModelRelationship>

        /// <summary>
        /// SupportByVersion Excel, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.ModelRelationship> GetEnumerator();

        #endregion
    }
}
