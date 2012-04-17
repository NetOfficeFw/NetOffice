using System;
using LateBindingApi.Core;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcPrintOrientation
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acPRORPortrait = 1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acPRORLandscape = 2
	}
}