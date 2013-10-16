using System;
using NetOffice;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Access", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcWebBrowserState
	{
		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acUnintialized = 0,

		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acLoading = 1,

		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acLoaded = 2,

		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acInteractive = 3,

		 /// <summary>
		 /// SupportByVersion Access 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 14,15)]
		 acComplete = 4
	}
}