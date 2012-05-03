using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlMousePointer
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerDefault = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerArrow = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerCross = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerIBeam = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerSizeNESW = 6,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerSizeNS = 7,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerSizeNWSE = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerSizeWE = 9,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerUpArrow = 10,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerHourGlass = 11,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerNoDrop = 12,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerAppStarting = 13,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerHelp = 14,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerSizeAll = 15,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14)]
		 olMousePointerCustom = 99
	}
}