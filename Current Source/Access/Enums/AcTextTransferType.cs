using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcTextTransferType
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImportDelim = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImportFixed = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acExportDelim = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acExportFixed = 3,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acExportMerge = 4,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acLinkDelim = 5,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acLinkFixed = 6,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acImportHTML = 7,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acExportHTML = 8,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Access", 9,10,11,12,14)]
		 acLinkHTML = 9
	}
}