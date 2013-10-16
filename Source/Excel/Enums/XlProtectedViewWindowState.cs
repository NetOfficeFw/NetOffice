using System;
using NetOffice;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Excel", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlProtectedViewWindowState
	{
		 /// <summary>
		 /// SupportByVersion Excel 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Excel", 14,15)]
		 xlProtectedViewWindowNormal = 0,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Excel", 14,15)]
		 xlProtectedViewWindowMinimized = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Excel", 14,15)]
		 xlProtectedViewWindowMaximized = 2
	}
}