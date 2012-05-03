using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 14
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjServerPage
	{
		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageApprovals = 0,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageResourceCenter = 1,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageProjectCenter = 2,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPagePortfolioAnalyzer = 3,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageStatusReports = 4,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPagePermissions = 5,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageProjectDetails = 6,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageDocuments = 7,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageIssues = 8,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageRisks = 9,

		 /// <summary>
		 /// SupportByVersion MSProject 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("MSProject", 14)]
		 pjServerPageWorkspace = 10
	}
}