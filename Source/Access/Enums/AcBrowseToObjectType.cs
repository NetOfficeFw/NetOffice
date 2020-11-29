using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcBrowseToObjectType"/> </remarks>
	[SupportByVersion("Access", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcBrowseToObjectType
	{
		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acBrowseToForm = 2,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acBrowseToReport = 3
	}
}