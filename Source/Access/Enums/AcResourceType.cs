using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcResourceType"/> </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcResourceType
	{
		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acResourceTheme = 0,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acResourceImage = 1
	}
}