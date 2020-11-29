using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OlEditorType"/> </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlEditorType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olEditorText = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olEditorHTML = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olEditorRTF = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		 olEditorWord = 4
	}
}