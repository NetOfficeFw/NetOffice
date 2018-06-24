using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIODATARECORDSETS 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Visio", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPVISIODATARECORDSETS : ICOMObject, IEnumerableProvider<NetOffice.VisioApi.IVDataRecordset>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.VisioApi.IVDataRecordset this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVDataRecordset get_ItemFromID(Int32 iD);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="iD">Int32 iD</param>
		[SupportByVersion("Visio", 12,14,15,16), Redirect("get_ItemFromID")]
		NetOffice.VisioApi.IVDataRecordset ItemFromID(Int32 iD);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="connectionIDOrString">object connectionIDOrString</param>
		/// <param name="commandString">string commandString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDataRecordset Add(object connectionIDOrString, string commandString, Int32 addOptions, object name);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="connectionIDOrString">object connectionIDOrString</param>
		/// <param name="commandString">string commandString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.IVDataRecordset Add(object connectionIDOrString, string commandString, Int32 addOptions);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="xMLString">string xMLString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDataRecordset AddFromXML(string xMLString, Int32 addOptions, object name);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="xMLString">string xMLString</param>
		/// <param name="addOptions">Int32 addOptions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.IVDataRecordset AddFromXML(string xMLString, Int32 addOptions);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="addOptions">Int32 addOptions</param>
		/// <param name="name">optional string Name = </param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDataRecordset AddFromConnectionFile(string fileName, Int32 addOptions, object name);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="addOptions">Int32 addOptions</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.IVDataRecordset AddFromConnectionFile(string fileName, Int32 addOptions);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataErrorCode">Int32 dataErrorCode</param>
		/// <param name="dataErrorDescription">string dataErrorDescription</param>
		/// <param name="recordsetID">Int32 recordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void GetLastDataError(out Int32 dataErrorCode, out string dataErrorDescription, out Int32 recordsetID);

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVDataRecordset>

        /// <summary>
        /// SupportByVersion Visio, 12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.VisioApi.IVDataRecordset> GetEnumerator();

        #endregion
    }
}
