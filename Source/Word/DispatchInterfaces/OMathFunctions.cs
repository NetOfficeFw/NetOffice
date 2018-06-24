using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface OMathFunctions 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845910.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("497142A4-16FD-42C6-BC58-15D89345FC21")]
	public interface OMathFunctions : ICOMObject, IEnumerableProvider<NetOffice.WordApi.OMathFunction>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823229.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838138.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837013.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836429.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.OMathFunction this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192335.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdOMathFunctionType type</param>
		/// <param name="numArgs">optional object numArgs</param>
		/// <param name="numCols">optional object numCols</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.OMathFunction Add(NetOffice.WordApi.Range range, NetOffice.WordApi.Enums.WdOMathFunctionType type, object numArgs, object numCols);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192335.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdOMathFunctionType type</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.OMathFunction Add(NetOffice.WordApi.Range range, NetOffice.WordApi.Enums.WdOMathFunctionType type);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192335.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdOMathFunctionType type</param>
		/// <param name="numArgs">optional object numArgs</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.OMathFunction Add(NetOffice.WordApi.Range range, NetOffice.WordApi.Enums.WdOMathFunctionType type, object numArgs);

        #endregion


        #region IEnumerable<NetOffice.WordApi.OMathFunction>

        /// <summary>
        /// SupportByVersion Word, 12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.OMathFunction> GetEnumerator();

        #endregion
    }
}
