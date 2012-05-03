using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlFormatNumber
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberAllDigits = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberTruncated = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumber1Decimal = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumber2Decimal = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberScientific = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberComputer1 = 6,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberComputer2 = 7,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberComputer3 = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olFormatNumberRaw = 9
	}
}