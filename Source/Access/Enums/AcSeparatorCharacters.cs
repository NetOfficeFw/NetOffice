using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Access", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcSeparatorCharacters
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acSeparatorCharactersSystemSeparator = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acSeparatorCharactersNewLine = 1,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acSeparatorCharactersSemiColon = 2,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15)]
		 acSeparatorCharactersComma = 3
	}
}