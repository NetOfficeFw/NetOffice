using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjGanttBarLink
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjNoGanttBarLinks = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjToTop = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14)]
		 pjToEnd = 2
	}
}