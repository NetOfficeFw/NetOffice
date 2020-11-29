﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcPrintItemLayout"/> </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcPrintItemLayout
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1953</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRHorizontalColumnLayout = 1953,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1954</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRVerticalColumnLayout = 1954
	}
}