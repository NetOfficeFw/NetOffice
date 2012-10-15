using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjReportLayoutTemplateId
	{
		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjReportLayoutTitleOnly = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjReportLayoutTitleAndChart = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjReportLayoutTitleAndTable = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 15)]
		 pjReportLayoutComparison = 3
	}
}