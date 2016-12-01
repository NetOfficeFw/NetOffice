using System;
using NetOffice;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 12, 14, 5.3
	 /// </summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum vbext_ComponentType
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_ct_StdModule = 1,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_ct_ClassModule = 2,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_ct_MSForm = 3,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_ct_ActiveXDesigner = 11,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_ct_Document = 100
	}
}