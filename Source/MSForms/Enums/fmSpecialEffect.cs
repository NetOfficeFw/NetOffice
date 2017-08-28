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
	public enum fmSpecialEffect
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmSpecialEffectFlat = 0,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmSpecialEffectRaised = 1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmSpecialEffectSunken = 2,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmSpecialEffectEtched = 3,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmSpecialEffectBump = 6
	}
}