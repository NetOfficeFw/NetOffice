using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSHTMLApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSHTML 4
	 /// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsEnum)]
	public enum _BEHAVIOR_LAYOUT_INFO
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORLAYOUTINFO_FULLDELEGATION = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORLAYOUTINFO_MODIFYNATURAL = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIORLAYOUTINFO_MAPSIZE = 4,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersion("MSHTML", 4)]
		 BEHAVIOR_LAYOUT_INFO_Max = 2147483647
	}
}