using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Access", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcBrowseToObjectType
	{
		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acBrowseToForm = 2,

		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acBrowseToReport = 3
	}
}