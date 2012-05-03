using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcImeSentenceMode
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImeSentenceModePhrasePredict = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImeSentenceModePluralClause = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImeSentenceModeConversation = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImeSentenceModeNone = 3
	}
}