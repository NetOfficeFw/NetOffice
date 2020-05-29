using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 5.3, 12, 14
	 /// </summary>
	[SupportByVersion("VBIDE", 5.3, 12, 14)]
	[EntityType(EntityType.IsEnum)]
	public enum vbext_RefKind
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_rk_TypeLib = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_rk_Project = 1
	}
}