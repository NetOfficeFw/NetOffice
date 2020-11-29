using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcNewDatabaseFormat"/> </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcNewDatabaseFormat
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acNewDatabaseFormatUserDefault = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acNewDatabaseFormatAccess2000 = 9,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acNewDatabaseFormatAccess2002 = 10,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acNewDatabaseFormatAccess2007 = 12
	}
}