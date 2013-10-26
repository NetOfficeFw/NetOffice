using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisRefreshSettings
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visRefreshOverwriteAll = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visRefreshNoReconcilationUI = 2
	}
}