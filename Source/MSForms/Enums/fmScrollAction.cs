using System;
using NetOffice;
namespace NetOffice.MSFormsApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSForms 2
	 /// </summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum fmScrollAction
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionNoChange = 0,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionLineUp = 1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionLineDown = 2,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionPageUp = 3,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionPageDown = 4,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionBegin = 5,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionEnd = 6,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 _fmScrollActionAbsoluteChange = 7,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionPropertyChange = 8,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionControlRequest = 9,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("MSForms", 2)]
		 fmScrollActionFocusRequest = 10
	}
}