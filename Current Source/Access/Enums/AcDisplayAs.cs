using System;
using LateBindingApi.Core;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Access", 12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum AcDisplayAs
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDisplayAsImageIcon = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDisplayAsIcon = 1,

		 /// <summary>
		 /// SupportByVersion Access 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Access", 12,14)]
		 acDisplayAsPaperclip = 2
	}
}