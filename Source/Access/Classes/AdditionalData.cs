using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass AdditionalData
	/// SupportByVersion Access, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191895.aspx </remarks>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("DBC51761-A8ED-11D3-A0DD-00C04F68712B")]
 	public interface AdditionalData : _AdditionalData
	{

	}
}
