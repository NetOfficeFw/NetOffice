using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVGraphicItem 
	/// SupportByVersion Visio, 12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0735-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.GraphicItem))]
    public interface IVGraphicItem : ICOMObject
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
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisObjectTypes ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 Index { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		string Description { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisGraphicItemTypes Type { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster DataGraphic { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		string Tag { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisGraphicPositionHorizontal HorizontalPosition { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisGraphicPositionVertical VerticalPosition { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool UseDataGraphicPosition { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="field">NetOffice.VisioApi.Enums.VisGraphicField field</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void GetExpression(out NetOffice.VisioApi.Enums.VisGraphicField field, out string expression);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="field">NetOffice.VisioApi.Enums.VisGraphicField field</param>
		/// <param name="expression">string expression</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void SetExpression(NetOffice.VisioApi.Enums.VisGraphicField field, string expression);

		#endregion
	}
}
