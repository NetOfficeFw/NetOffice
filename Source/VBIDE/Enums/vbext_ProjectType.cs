using System;
using NetOffice;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 12, 14, 5.3
	 /// </summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum vbext_ProjectType
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pt_HostProject = 100,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pt_StandAlone = 101
	}
}