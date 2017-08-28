using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823069.aspx </remarks>
	[SupportByVersion("Access", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcTransformXMLScriptOption
	{
		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 11,12,14,15,16)]
		 acEnableScript = 0,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 11,12,14,15,16)]
		 acPromptScript = 1,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 11,12,14,15,16)]
		 acDisableScript = 2
	}
}