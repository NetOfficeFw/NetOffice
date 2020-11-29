using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoSmartArtNodePosition"/> </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoSmartArtNodePosition
	{
		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoSmartArtNodeDefault = 1,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoSmartArtNodeAfter = 2,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoSmartArtNodeBefore = 3,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoSmartArtNodeAbove = 4,

		 /// <summary>
		 /// SupportByVersion Office 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 14,15,16)]
		 msoSmartArtNodeBelow = 5
	}
}