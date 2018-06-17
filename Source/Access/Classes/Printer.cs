using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass Printer 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837177.aspx </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("DBC5175E-A8ED-11D3-A0DD-00C04F68712B")]
 	public interface Printer : _Printer
	{

	}
}
