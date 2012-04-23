using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoPermission
	{
		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionView = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionRead = 1,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionEdit = 2,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionSave = 4,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionExtract = 8,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionChange = 15,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionPrint = 16,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionObjModel = 32,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoPermissionFullControl = 64,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoPermissionAllCommon = 127
	}
}