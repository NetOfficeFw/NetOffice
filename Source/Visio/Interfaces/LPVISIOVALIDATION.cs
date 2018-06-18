using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOVALIDATION 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPVISIOVALIDATION : ICOMObject
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
		DateTime LastValidatedDate { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool ShowIgnoredIssues { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVValidationRuleSets RuleSets { get; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVValidationIssues Issues { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="ruleSet">optional NetOffice.VisioApi.IVValidationRuleSet ruleSet</param>
		/// <param name="flags">optional NetOffice.VisioApi.Enums.VisValidationFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void Validate(object ruleSet, object flags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void Validate();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="ruleSet">optional NetOffice.VisioApi.IVValidationRuleSet ruleSet</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void Validate(object ruleSet);

		#endregion
	}
}
