using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 11)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjReportLayoutTemplateId
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjReportLayoutTitleOnly = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjReportLayoutTitleAndChart = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjReportLayoutTitleAndTable = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 11
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 11)]
		 pjReportLayoutComparison = 3
	}
}