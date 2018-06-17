using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// CoClass ReturnVars 
	/// SupportByVersion Access, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192693.aspx </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("B06693E3-385D-4E70-923E-4FAB6D14EE15")]
    [CoClassSource(typeof(NetOffice.AccessApi.ReturnVars))]
    public interface ReturnVars : _ReturnVars
	{

	}
}
