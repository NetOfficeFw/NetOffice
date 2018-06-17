using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IReturnEffect 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("8A683C91-BA84-11CF-8110-00A0C9030074")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.ReturnEffect))]
    public interface IReturnEffect : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmDropEffect Value { get; set; }

		#endregion

	}
}
