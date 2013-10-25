using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823069.aspx </remarks>
	[SupportByVersionAttribute("Access", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcTransformXMLScriptOption
	{
		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acEnableScript = 0,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acPromptScript = 1,

		 /// <summary>
		 /// SupportByVersion Access 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 11,12,14,15)]
		 acDisableScript = 2
	}
}