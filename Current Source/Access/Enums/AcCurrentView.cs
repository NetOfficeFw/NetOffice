using System;
using LateBindingApi.Core;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcCurrentView
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acCurViewDesign = 0,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acCurViewFormBrowse = 1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acCurViewDatasheet = 2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acCurViewPivotTable = 3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acCurViewPivotChart = 4,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acCurViewPreview = 5,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acCurViewReportBrowse = 6,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acCurViewLayout = 7
	}
}