﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlSaveAsType"/> </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlSaveAsType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTXT = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olRTF = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olTemplate = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olMSG = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olDoc = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olHTML = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olVCard = 6,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olVCal = 7,

		 /// <summary>
		 /// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Outlook", 10,11,12,14,15,16)]
		 olICal = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Outlook", 11,12,14,15,16)]
		 olMSGUnicode = 9,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olMHTML = 10
	}
}