using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 5.3, 12, 14, 15
	 /// </summary>
	[SupportByVersion("VBIDE", 5.3,12,14,15)]
	[EntityType(EntityType.IsEnum)]
	public enum vbext_ComponentType
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("VBIDE", 5.3,12,14,15)]
		 vbext_ct_StdModule = 1,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("VBIDE", 5.3,12,14,15)]
		 vbext_ct_ClassModule = 2,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("VBIDE", 5.3,12,14,15)]
		 vbext_ct_MSForm = 3,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("VBIDE", 5.3,12,14,15)]
		 vbext_ct_ActiveXDesigner = 11,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14, 15
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("VBIDE", 5.3,12,14,15)]
		 vbext_ct_Document = 100
	}
}