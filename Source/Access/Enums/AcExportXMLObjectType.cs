﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcExportXMLObjectType"/> </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcExportXMLObjectType
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportTable = 0,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportQuery = 1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportForm = 2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportReport = 3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportServerView = 7,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportStoredProcedure = 9,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acExportFunction = 10
	}
}