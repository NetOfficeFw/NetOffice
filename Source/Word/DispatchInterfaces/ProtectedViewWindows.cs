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
	/// DispatchInterface ProtectedViewWindows 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197163.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Word", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9")]
	public interface ProtectedViewWindows : ICOMObject, IEnumerableProvider<NetOffice.WordApi.ProtectedViewWindow>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196633.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821613.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822106.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840128.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.WordApi.ProtectedViewWindow this[object index] { get; }

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument, object visible, object openAndRepair);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ProtectedViewWindow Open(object fileName);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument);

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193715.aspx </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument, object visible);

        #endregion

        #region IEnumerable<NetOffice.WordApi.ProtectedViewWindow>

        /// <summary>
        /// SupportByVersion Word, 14,15,16
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        new IEnumerator<NetOffice.WordApi.ProtectedViewWindow> GetEnumerator();

        #endregion
    }
}
