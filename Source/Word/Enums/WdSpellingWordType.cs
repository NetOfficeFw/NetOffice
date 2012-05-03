using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdSpellingWordType
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSpellword = 0,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdWildcard = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdAnagram = 2
	}
}