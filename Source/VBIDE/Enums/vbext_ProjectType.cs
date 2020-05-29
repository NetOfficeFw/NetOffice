﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 5.3, 12, 14
	 /// </summary>
	[SupportByVersion("VBIDE", 5.3, 12, 14)]
	[EntityType(EntityType.IsEnum)]
	public enum vbext_ProjectType
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_pt_HostProject = 100,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_pt_StandAlone = 101
	}
}