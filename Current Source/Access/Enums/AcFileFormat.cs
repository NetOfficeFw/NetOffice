using System;
using LateBindingApi.Core;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcFileFormat
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFileFormatAccess2 = 2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFileFormatAccess95 = 7,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFileFormatAccess97 = 8,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFileFormatAccess2000 = 9,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Access", 10,11,12,14)]
		 acFileFormatAccess2002 = 10,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acFileFormatAccess2007 = 12
	}
}