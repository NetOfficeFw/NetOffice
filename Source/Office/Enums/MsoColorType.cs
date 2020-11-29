using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoColorType"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoColorType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoColorTypeMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoColorTypeRGB = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoColorTypeScheme = 2,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoColorTypeCMYK = 3,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoColorTypeCMS = 4,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoColorTypeInk = 5
	}
}