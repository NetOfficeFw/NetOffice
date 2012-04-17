using System;
using LateBindingApi.Core;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum XlChartGallery
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlBuiltIn = 21,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlUserDefined = 22,

		 /// <summary>
		 /// SupportByVersion PowerPoint 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("PowerPoint", 14)]
		 xlAnyGallery = 23
	}
}