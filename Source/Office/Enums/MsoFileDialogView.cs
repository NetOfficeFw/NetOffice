using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff865533.aspx </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoFileDialogView
	{
		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewList = 1,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewDetails = 2,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewProperties = 3,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewPreview = 4,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewThumbnail = 5,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewLargeIcons = 6,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewSmallIcons = 7,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoFileDialogViewWebView = 8,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoFileDialogViewTiles = 9
	}
}