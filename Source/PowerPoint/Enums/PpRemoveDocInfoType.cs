using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpRemoveDocInfoType
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIComments = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIRemovePersonalInformation = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIDocumentProperties = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIDocumentWorkspace = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIInkAnnotations = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIPublishPath = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIDocumentServerProperties = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIDocumentManagementPolicy = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIContentType = 16,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDISlideUpdateInformation = 17,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppRDIAll = 99
	}
}