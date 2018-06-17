using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IReturnSingle 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("8A683C90-BA84-11CF-8110-00A0C9030074")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.ReturnSingle))]
    public interface IReturnSingle : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single Value { get; set; }

		#endregion

	}
}
