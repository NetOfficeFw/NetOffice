using System;
using NetOffice;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 12, 14, 5.3
	 /// </summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum vbext_ProcKind
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pk_Proc = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pk_Let = 1,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pk_Set = 2,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		 vbext_pk_Get = 3
	}
}