using System;
using NetOffice;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 12, 14, 5.3
	 /// </summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum vbext_ProjectProtection
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pp_none = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pp_locked = 1
	}
}