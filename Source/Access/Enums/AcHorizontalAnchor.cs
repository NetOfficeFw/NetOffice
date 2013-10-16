using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Access", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcHorizontalAnchor
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acHorizontalAnchorLeft = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acHorizontalAnchorRight = 1,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acHorizontalAnchorBoth = 2
	}
}