using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/overview/Library-Reference/msolinefilltype-enumeration-office"/> </remarks>
	[SupportByVersion("Office", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoLineFillType
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillNone = 0,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillSolid = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillPatterned = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillGradient = 3,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillTextured = 4,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillBackground = 5,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoLineFillPicture = 6
	}
}