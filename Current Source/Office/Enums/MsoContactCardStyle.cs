using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoContactCardStyle
	{
		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoContactCardHover = 0,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoContactCardFull = 1
	}
}