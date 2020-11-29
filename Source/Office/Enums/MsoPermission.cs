﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoPermission"/> </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoPermission
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionView = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionRead = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionEdit = 2,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionSave = 4,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionExtract = 8,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionChange = 15,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionPrint = 16,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionObjModel = 32,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoPermissionFullControl = 64,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPermissionAllCommon = 127
	}
}