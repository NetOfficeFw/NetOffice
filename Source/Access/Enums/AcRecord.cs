using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcRecord
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acPrevious = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acNext = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acFirst = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acLast = 3,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acGoTo = 4,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acNewRec = 5
	}
}