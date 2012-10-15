using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjWeekday
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjSunday = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjMonday = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjTuesday = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjWednesday = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjThursday = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjFriday = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 11,12,14,15)]
		 pjSaturday = 7
	}
}