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
	/// DispatchInterface BuildingBlockEntries 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835133.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Word", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("39709229-56A0-4E29-9112-B31DD067EBFD")]
	public interface BuildingBlockEntries : ICOMObject, IEnumerableProvider<NetOffice.WordApi.BuildingBlock>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840334.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194454.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836710.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838105.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.BuildingBlock this[object index] { get; }

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845259.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdBuildingBlockTypes type</param>
		/// <param name="category">string category</param>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="description">optional object description</param>
		/// <param name="insertOptions">optional NetOffice.WordApi.Enums.WdDocPartInsertOptions InsertOptions = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.BuildingBlock Add(string name, NetOffice.WordApi.Enums.WdBuildingBlockTypes type, string category, NetOffice.WordApi.Range range, object description, object insertOptions);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845259.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdBuildingBlockTypes type</param>
		/// <param name="category">string category</param>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.BuildingBlock Add(string name, NetOffice.WordApi.Enums.WdBuildingBlockTypes type, string category, NetOffice.WordApi.Range range);

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845259.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.WordApi.Enums.WdBuildingBlockTypes type</param>
		/// <param name="category">string category</param>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		NetOffice.WordApi.BuildingBlock Add(string name, NetOffice.WordApi.Enums.WdBuildingBlockTypes type, string category, NetOffice.WordApi.Range range, object description);

        #endregion

        #region IEnumerable<NetOffice.WordApi.BuildingBlock>

        /// <summary>
        /// SupportByVersion Word, 12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.WordApi.BuildingBlock> GetEnumerator();

        #endregion
    }
}
