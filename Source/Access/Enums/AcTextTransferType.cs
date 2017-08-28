using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194227.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcTextTransferType
	{
		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acImportDelim = 0,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acImportFixed = 1,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acExportDelim = 2,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acExportFixed = 3,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acExportMerge = 4,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acLinkDelim = 5,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acLinkFixed = 6,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acImportHTML = 7,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acExportHTML = 8,

		 /// <summary>
		 /// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		 acLinkHTML = 9
	}
}