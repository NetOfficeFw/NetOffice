using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcDefView"/> </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcDefView
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acDefViewSingle = 0,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acDefViewContinuous = 1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acDefViewDatasheet = 2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acDefViewPivotTable = 3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acDefViewPivotChart = 4,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acDefViewSplitForm = 5
	}
}