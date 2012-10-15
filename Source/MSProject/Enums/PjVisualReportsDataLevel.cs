using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjVisualReportsDataLevel
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjLevelYears = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjLevelQuarters = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjLevelMonths = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjLevelWeeks = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjLevelDays = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjLevelAutomatic = 5
	}
}