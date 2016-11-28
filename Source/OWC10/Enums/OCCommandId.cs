using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OCCommandId
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1007</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandAbout = 1007,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandUndo = 1000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandCut = 1001,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1002</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandCopy = 1002,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1003</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandPaste = 1003,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1005</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandProperties = 1005,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1006</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandHelp = 1006,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1004</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandExport = 1004,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2000</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandSortAsc = 2000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2031</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandSortDesc = 2031,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1010</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandChooser = 1010,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1017</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandAutoFilter = 1017,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1016</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandAutoCalc = 1016,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1013</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandCollapse = 1013,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1012</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandExpand = 1012,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1014</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ocCommandRefresh = 1014
	}
}