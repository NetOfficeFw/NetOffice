using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdInternationalIndex
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdListSeparator = 17,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDecimalSeparator = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdThousandsSeparator = 19,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdCurrencyCode = 20,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wd24HourClock = 21,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdInternationalAM = 22,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdInternationalPM = 23,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdTimeSeparator = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDateSeparator = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdProductLanguageID = 26
	}
}