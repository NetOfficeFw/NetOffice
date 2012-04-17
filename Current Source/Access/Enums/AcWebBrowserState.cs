using System;
using LateBindingApi.Core;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcWebBrowserState
	{
		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acUnintialized = 0,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acLoading = 1,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acLoaded = 2,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acInteractive = 3,

		 /// <summary>
		 /// SupportByVersion Access 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Access", 14)]
		 acComplete = 4
	}
}