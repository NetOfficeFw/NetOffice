using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass LocalVar 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("8357BB51-95A2-4043-A040-2825FACEF50D")]
 	public interface LocalVar : _LocalVar
	{

	}
}
