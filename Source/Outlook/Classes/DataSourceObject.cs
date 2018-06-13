using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// CoClass DataSourceObject 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsCoClass)]
	[TypeId("0006F02C-0000-0000-C000-000000000046")]
 	public interface DataSourceObject : DDataSourceObject
	{

	}
}
