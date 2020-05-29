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
	public enum vbext_VBAMode
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_vm_Run = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_vm_Break = 1,

		 /// <summary>
		 /// SupportByVersion VBIDE 5.3, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("VBIDE", 5.3, 12, 14)]
		 vbext_vm_Design = 2
	}
}