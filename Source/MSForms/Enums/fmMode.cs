using System;
using NetOffice;
namespace NetOffice.MSFormsApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSForms 2
	 /// </summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum fmMode
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmModeInherit = -2,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmModeOn = -1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmModeOff = 0
	}
}