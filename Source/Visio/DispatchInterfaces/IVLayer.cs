using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVLayer 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0712-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Layer))]
    public interface IVLayer : ICOMObject
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
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVPage Page { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster Master { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Index { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Row { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVCell get_CellsC(Int16 column);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsC
		/// </summary>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsC")]
		NetOffice.VisioApi.IVCell CellsC(Int16 column);

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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Stat { get; }

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string NameU { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="fPresMems">Int16 fPresMems</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Add(NetOffice.VisioApi.IVShape sheetObject, Int16 fPresMems);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="fPresMems">Int16 fPresMems</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Remove(NetOffice.VisioApi.IVShape sheetObject, Int16 fPresMems);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fDeleteShapes">Int16 fDeleteShapes</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Delete(Int16 fDeleteShapes);

		#endregion
	}
}
