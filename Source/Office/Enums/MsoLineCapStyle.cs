using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229980.aspx </remarks>
	[SupportByVersionAttribute("Office", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoLineCapStyle
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineCapMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineCapSquare = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineCapRound = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineCapFlat = 3
	}
}