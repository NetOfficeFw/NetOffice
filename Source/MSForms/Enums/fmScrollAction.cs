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
	public enum fmScrollAction
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionNoChange = 0,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionLineUp = 1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionLineDown = 2,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionPageUp = 3,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionPageDown = 4,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionBegin = 5,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionEnd = 6,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSForms", 2)]
		 _fmScrollActionAbsoluteChange = 7,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionPropertyChange = 8,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionControlRequest = 9,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmScrollActionFocusRequest = 10
	}
}