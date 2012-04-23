using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoFlipCmd
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoFlipHorizontal = 0,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoFlipVertical = 1
	}
}