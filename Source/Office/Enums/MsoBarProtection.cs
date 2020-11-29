using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoBarProtection"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoBarProtection
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoProtection = 0,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoCustomize = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoResize = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoMove = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoChangeVisible = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoChangeDock = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoVerticalDock = 32,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBarNoHorizontalDock = 64
	}
}