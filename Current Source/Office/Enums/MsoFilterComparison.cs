using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoFilterComparison
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonEqual = 0,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonNotEqual = 1,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonLessThan = 2,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonGreaterThan = 3,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonLessThanEqual = 4,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonGreaterThanEqual = 5,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonIsBlank = 6,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonIsNotBlank = 7,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonContains = 8,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoFilterComparisonNotContains = 9
	}
}