using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoCustomXMLNodeType"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoCustomXMLNodeType
	{
		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeElement = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeAttribute = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeText = 3,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeCData = 4,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeProcessingInstruction = 7,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeComment = 8,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoCustomXMLNodeDocument = 9
	}
}