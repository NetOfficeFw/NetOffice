using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768606(v=office.14).aspx </remarks>
	[SupportByVersionAttribute("Visio", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisBuiltInStencilTypes
	{
		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visBuiltInStencilBackgrounds = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visBuiltInStencilBorders = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visBuiltInStencilContainers = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visBuiltInStencilCallouts = 3,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visBuiltInStencilLegends = 4
	}
}