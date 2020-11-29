using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ExcelApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Excel 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.XlProtectedViewCloseReason"/> </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum XlProtectedViewCloseReason
	{
		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlProtectedViewCloseNormal = 0,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlProtectedViewCloseEdit = 1,

		 /// <summary>
		 /// SupportByVersion Excel 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Excel", 14,15,16)]
		 xlProtectedViewCloseForced = 2
	}
}