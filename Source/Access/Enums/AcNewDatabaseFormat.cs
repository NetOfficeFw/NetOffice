using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198328.aspx </remarks>
	[SupportByVersionAttribute("Access", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcNewDatabaseFormat
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15,16)]
		 acNewDatabaseFormatUserDefault = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15,16)]
		 acNewDatabaseFormatAccess2000 = 9,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15,16)]
		 acNewDatabaseFormatAccess2002 = 10,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Access", 12,14,15,16)]
		 acNewDatabaseFormatAccess2007 = 12
	}
}