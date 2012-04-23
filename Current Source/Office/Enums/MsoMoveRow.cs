using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoMoveRow
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoMoveRowFirst = -4,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoMoveRowPrev = -3,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoMoveRowNext = -2,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoMoveRowNbr = -1
	}
}