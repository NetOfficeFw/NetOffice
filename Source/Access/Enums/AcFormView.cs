using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcFormView
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acNormal = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acDesign = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acPreview = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acFormDS = 3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFormPivotTable = 4,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFormPivotChart = 5,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acLayout = 6
	}
}