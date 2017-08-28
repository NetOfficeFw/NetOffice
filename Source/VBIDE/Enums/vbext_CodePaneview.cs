using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VBIDEApi.Enums
{
	 /// <summary>
	 /// SupportByVersion VBIDE 12, 14, 5.3
	 /// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsEnum)]
	public enum vbext_CodePaneview
	{
		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_cv_ProcedureView = 0,

		 /// <summary>
		 /// SupportByVersion VBIDE 12, 14, 5.3
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("VBIDE", 12,14,5.3)]
		 vbext_cv_FullModuleView = 1
	}
}