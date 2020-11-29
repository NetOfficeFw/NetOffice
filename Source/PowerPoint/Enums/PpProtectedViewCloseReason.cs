using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.ppprotectedviewclosereason"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum PpProtectedViewCloseReason
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 ppProtectedViewCloseNormal = 0,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 ppProtectedViewCloseEdit = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("PowerPoint", 14,15,16)]
		 ppProtectedViewCloseForced = 2
	}
}