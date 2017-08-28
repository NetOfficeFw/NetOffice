using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSFormsApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSForms 2
	 /// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsEnum)]
	public enum fmMode
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmModeInherit = -2,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmModeOn = -1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmModeOff = 0
	}
}