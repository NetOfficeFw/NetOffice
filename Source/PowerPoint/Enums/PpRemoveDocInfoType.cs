using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745305.aspx </remarks>
	[SupportByVersionAttribute("PowerPoint", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpRemoveDocInfoType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIComments = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIRemovePersonalInformation = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIDocumentProperties = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIDocumentWorkspace = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIInkAnnotations = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIPublishPath = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIDocumentServerProperties = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIDocumentManagementPolicy = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIContentType = 16,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDISlideUpdateInformation = 17,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14, 15
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14,15)]
		 ppRDIAll = 99
	}
}