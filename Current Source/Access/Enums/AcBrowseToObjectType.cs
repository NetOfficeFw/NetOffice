using System;
using LateBindingApi.Core;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcBrowseToObjectType
	{
		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acBrowseToForm = 2,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acBrowseToReport = 3
	}
}