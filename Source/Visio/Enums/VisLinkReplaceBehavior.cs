using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisLinkReplaceBehavior
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visLinkReplaceNever = 0,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visLinkReplaceAlways = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visLinkReplacePrompt = 2
	}
}