using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Dummy 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("8B06E320-B23C-11CF-89A8-00A0C9054129")]
    [CoClassSource(typeof(NetOffice.AccessApi.Class))]
	public interface _Dummy : ICOMObject
	{
	}
}
