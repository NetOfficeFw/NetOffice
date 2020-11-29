using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcQuitOption"/> </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcQuitOption
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acQuitPrompt = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acQuitSaveAll = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acQuitSaveNone = 2
	}
}