using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoBarProtection
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoProtection = 0,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoCustomize = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoResize = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoMove = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoChangeVisible = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoChangeDock = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoVerticalDock = 32,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoBarNoHorizontalDock = 64
	}
}