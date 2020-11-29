using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/overview/Library-Reference/msolinecapstyle-enumeration-office"/> </remarks>
	[SupportByVersion("Office", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoLineCapStyle
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineCapMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineCapSquare = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineCapRound = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineCapFlat = 3
	}
}