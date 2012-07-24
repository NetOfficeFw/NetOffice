using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoLineCapStyle
	{
		 /// <summary>
		 /// SupportByVersion Office 15
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 15)]
		 msoLineCapMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 15)]
		 msoLineCapSquare = 1,

		 /// <summary>
		 /// SupportByVersion Office 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 15)]
		 msoLineCapRound = 2,

		 /// <summary>
		 /// SupportByVersion Office 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 15)]
		 msoLineCapFlat = 3
	}
}