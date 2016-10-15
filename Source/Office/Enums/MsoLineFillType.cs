using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228403.aspx </remarks>
	[SupportByVersionAttribute("Office", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoLineFillType
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillNone = 0,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillSolid = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillPatterned = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillGradient = 3,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillTextured = 4,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillBackground = 5,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoLineFillPicture = 6
	}
}