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
	public enum fmLayoutEffect
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmLayoutEffectNone = 0,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmLayoutEffectInitiate = 1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSForms", 2)]
		 _fmLayoutEffectRespond = 2
	}
}