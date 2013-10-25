using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860880.aspx </remarks>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoEncoding
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>874</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingThai = 874,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>932</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingJapaneseShiftJIS = 932,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>936</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingSimplifiedChineseGBK = 936,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>949</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingKorean = 949,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>950</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTraditionalChineseBig5 = 950,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingUnicodeLittleEndian = 1200,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1201</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingUnicodeBigEndian = 1201,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1250</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingCentralEuropean = 1250,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1251</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingCyrillic = 1251,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1252</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingWestern = 1252,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1253</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingGreek = 1253,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1254</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTurkish = 1254,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1255</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingHebrew = 1255,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1256</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingArabic = 1256,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1257</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingBaltic = 1257,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1258</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingVietnamese = 1258,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50001</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingAutoDetect = 50001,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50932</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingJapaneseAutoDetect = 50932,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50936</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingSimplifiedChineseAutoDetect = 50936,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50949</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingKoreanAutoDetect = 50949,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50950</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTraditionalChineseAutoDetect = 50950,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51251</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingCyrillicAutoDetect = 51251,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51253</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingGreekAutoDetect = 51253,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51256</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingArabicAutoDetect = 51256,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28591</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88591Latin1 = 28591,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28592</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88592CentralEurope = 28592,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28593</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88593Latin3 = 28593,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28594</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88594Baltic = 28594,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28595</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88595Cyrillic = 28595,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28596</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88596Arabic = 28596,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28597</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88597Greek = 28597,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28598</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88598Hebrew = 28598,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28599</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO88599Turkish = 28599,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28605</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO885915Latin9 = 28605,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50220</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO2022JPNoHalfwidthKatakana = 50220,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50221</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO2022JPJISX02021984 = 50221,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50222</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO2022JPJISX02011989 = 50222,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50225</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO2022KR = 50225,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50227</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO2022CNTraditionalChinese = 50227,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50229</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO2022CNSimplifiedChinese = 50229,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10000</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacRoman = 10000,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10001</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacJapanese = 10001,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10002</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacTraditionalChineseBig5 = 10002,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10003</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacKorean = 10003,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10004</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacArabic = 10004,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10005</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacHebrew = 10005,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10006</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacGreek1 = 10006,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10007</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacCyrillic = 10007,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10008</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacSimplifiedChineseGB2312 = 10008,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10010</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacRomania = 10010,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10017</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacUkraine = 10017,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10029</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacLatin2 = 10029,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10079</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacIcelandic = 10079,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10081</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacTurkish = 10081,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10082</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingMacCroatia = 10082,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICUSCanada = 37,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICInternational = 500,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>870</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICMultilingualROECELatin2 = 870,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>875</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICGreekModern = 875,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1026</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICTurkishLatin5 = 1026,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20273</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICGermany = 20273,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20277</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICDenmarkNorway = 20277,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20278</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICFinlandSweden = 20278,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20280</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICItaly = 20280,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20284</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICLatinAmericaSpain = 20284,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20285</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICUnitedKingdom = 20285,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20290</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICJapaneseKatakanaExtended = 20290,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20297</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICFrance = 20297,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20420</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICArabic = 20420,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20423</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICGreek = 20423,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20424</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICHebrew = 20424,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20833</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICKoreanExtended = 20833,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20838</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICThai = 20838,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20871</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICIcelandic = 20871,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20905</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICTurkish = 20905,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20880</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICRussian = 20880,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21025</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICSerbianBulgarian = 21025,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50930</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50931</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICUSCanadaandJapanese = 50931,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50933</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICKoreanExtendedAndKorean = 50933,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50935</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50937</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICUSCanadaAndTraditionalChinese = 50937,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50939</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>437</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMUnitedStates = 437,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>737</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMGreek437G = 737,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>775</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMBaltic = 775,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>850</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMMultilingualLatinI = 850,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>852</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMMultilingualLatinII = 852,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>855</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMCyrillic = 855,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>857</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMTurkish = 857,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>860</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMPortuguese = 860,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>861</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMIcelandic = 861,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>862</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMHebrew = 862,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>863</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMCanadianFrench = 863,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>864</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMArabic = 864,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>865</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMNordic = 865,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>866</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMCyrillicII = 866,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>869</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingOEMModernGreek = 869,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51932</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEUCJapanese = 51932,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51936</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEUCChineseSimplifiedChinese = 51936,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51949</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEUCKorean = 51949,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>51950</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEUCTaiwaneseTraditionalChinese = 51950,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>708</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingArabicASMO = 708,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>720</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingArabicTransparentASMO = 720,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingKoreanJohab = 1361,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20000</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTaiwanCNS = 20000,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20001</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTaiwanTCA = 20001,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20002</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTaiwanEten = 20002,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20003</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTaiwanIBM5550 = 20003,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20004</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTaiwanTeleText = 20004,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20005</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingTaiwanWang = 20005,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20105</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingIA5IRV = 20105,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20106</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingIA5German = 20106,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20107</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingIA5Swedish = 20107,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20108</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingIA5Norwegian = 20108,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20127</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingUSASCII = 20127,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20261</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingT61 = 20261,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20269</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingISO6937NonSpacingAccent = 20269,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20866</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingKOI8R = 20866,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21027</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingExtAlphaLowercase = 21027,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21866</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingKOI8U = 21866,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>29001</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingEuropa3 = 29001,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>52936</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingHZGBSimplifiedChinese = 52936,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65000</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingUTF7 = 65000,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>65001</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoEncodingUTF8 = 65001,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57002</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIDevanagari = 57002,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57003</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIBengali = 57003,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57004</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIITamil = 57004,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57005</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIITelugu = 57005,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57006</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIAssamese = 57006,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57007</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIOriya = 57007,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57008</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIKannada = 57008,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57009</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIMalayalam = 57009,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57010</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIGujarati = 57010,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>57011</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoEncodingISCIIPunjabi = 57011,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>38598</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15)]
		 msoEncodingISO88598HebrewLogical = 38598,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>54936</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14,15)]
		 msoEncodingSimplifiedChineseGB18030 = 54936
	}
}