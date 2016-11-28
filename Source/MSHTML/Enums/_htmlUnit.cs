using System;
using NetOffice;
namespace NetOffice.MSHTMLApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSHTML 4
	 /// </summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum _htmlUnit
	{
		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlUnitCharacter = 1,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlUnitWord = 2,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlUnitSentence = 3,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlUnitTextEdit = 6,

		 /// <summary>
		 /// SupportByVersion MSHTML 4
		 /// </summary>
		 /// <remarks>2147483647</remarks>
		 [SupportByVersionAttribute("MSHTML", 4)]
		 htmlUnit_Max = 2147483647
	}
}