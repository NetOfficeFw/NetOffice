using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcFileFormat"/> </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcFileFormat
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acFileFormatAccess2 = 2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acFileFormatAccess95 = 7,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acFileFormatAccess97 = 8,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acFileFormatAccess2000 = 9,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acFileFormatAccess2002 = 10,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acFileFormatAccess2007 = 12
	}
}