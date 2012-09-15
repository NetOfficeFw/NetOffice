using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 12, 3.6
	 /// </summary>
	[SupportByVersionAttribute("DAO", 12,3.6)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum CollatingOrderEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortNeutral = 1024,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1025</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortArabic = 1025,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1049</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortCyrillic = 1049,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1029</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortCzech = 1029,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1043</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortDutch = 1043,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1033</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortGeneral = 1033,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1032</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortGreek = 1032,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1037</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortHebrew = 1037,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1038</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortHungarian = 1038,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1039</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortIcelandic = 1039,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1030</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortNorwdan = 1030,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1033</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortPDXIntl = 1033,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1030</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortPDXNor = 1030,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1053</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortPDXSwe = 1053,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1045</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortPolish = 1045,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1034</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortSpanish = 1034,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1053</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortSwedFin = 1053,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1055</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortTurkish = 1055,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1041</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortJapanese = 1041,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>2052</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortChineseSimplified = 2052,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1028</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortChineseTraditional = 1028,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1042</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortKorean = 1042,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1054</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortThai = 1054,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>1060</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortSlovenian = 1060,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbSortUndefined = -1,

		 /// <summary>
		 /// SupportByVersion DAO 12
		 /// </summary>
		 /// <remarks>263185</remarks>
		 [SupportByVersionAttribute("DAO", 12)]
		 dbSortJapaneseRadicalStrokeCount = 263185,

		 /// <summary>
		 /// SupportByVersion DAO 12
		 /// </summary>
		 /// <remarks>1081</remarks>
		 [SupportByVersionAttribute("DAO", 12)]
		 dbSortHindi = 1081
	}
}