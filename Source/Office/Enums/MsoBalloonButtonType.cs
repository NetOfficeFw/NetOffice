using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227975.aspx </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoBalloonButtonType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-15</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonYesToAll = -15,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-14</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonOptions = -14,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-13</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonTips = -13,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-12</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonClose = -12,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-11</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonSnooze = -11,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-10</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonSearch = -10,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-9</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonIgnore = -9,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-8</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonAbort = -8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-7</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonRetry = -7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-6</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonNext = -6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-5</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonBack = -5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-4</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonNo = -4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-3</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonYes = -3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonCancel = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonOK = -1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoBalloonButtonNull = 0
	}
}