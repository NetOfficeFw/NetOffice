using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjCalendarType
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjGregorianCalendar = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjHijriCalendar = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjThaiCalendar = 7
	}
}