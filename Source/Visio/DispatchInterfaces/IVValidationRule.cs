using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVValidationRule 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D073E-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.ValidationRule))]
    public interface IVValidationRule : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		string NameU { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		string Category { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		string Description { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool Ignored { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		string FilterExpression { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRuleTargets TargetType { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		string TestExpression { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVValidationRuleSet RuleSet { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="targetPage">optional NetOffice.VisioApi.IVPage targetPage</param>
		/// <param name="targetShape">optional NetOffice.VisioApi.IVShape targetShape</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVValidationIssue AddIssue(object targetPage, object targetShape);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.IVValidationIssue AddIssue();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="targetPage">optional NetOffice.VisioApi.IVPage targetPage</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.IVValidationIssue AddIssue(object targetPage);

		#endregion
	}
}
