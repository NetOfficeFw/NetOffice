using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcSplitFormOrientation
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDatasheetOnTop = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDatasheetOnBottom = 1,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDatasheetOnLeft = 2,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDatasheetOnRight = 3
	}
}