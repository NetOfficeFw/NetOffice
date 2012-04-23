using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcHyperlinkPart
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acDisplayedValue = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acDisplayText = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acAddress = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acSubAddress = 3,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acScreenTip = 4,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acFullAddress = 5
	}
}