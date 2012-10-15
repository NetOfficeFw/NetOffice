using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjFileFormat
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjMPP = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTXT = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjCSV = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjXLS = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjMPT = 11,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjXLSX = 20,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjXLSB = 21,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjMPP9 = 22,

		 /// <summary>
		 /// SupportByVersion MSProject 14, 15
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("MSProject", 14,15)]
		 pjMPP12 = 23
	}
}