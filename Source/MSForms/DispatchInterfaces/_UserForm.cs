using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface _UserForm 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("04598FC8-866C-11CF-AB7C-00AA00C08FCF")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.UserForm))]
    public interface _UserForm : IOptionFrame
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 DrawBuffer { get; set; }

		#endregion

	}
}
