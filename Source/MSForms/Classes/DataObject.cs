using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// CoClass DataObject 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("1C3B4210-F441-11CE-B9EA-00AA006B1A69")]
 	public interface DataObject : IDataAutoWrapper
	{

	}
}
