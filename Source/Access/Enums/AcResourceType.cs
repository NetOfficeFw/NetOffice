using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcResourceType
	{
		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acResourceTheme = 0,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acResourceImage = 1
	}
}