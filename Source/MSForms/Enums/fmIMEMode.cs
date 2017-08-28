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
	public enum fmIMEMode
	{
		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeNoControl = 0,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeOn = 1,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeOff = 2,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeDisable = 3,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeHiragana = 4,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeKatakana = 5,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeKatakanaHalf = 6,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeAlphaFull = 7,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeAlpha = 8,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeHangulFull = 9,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeHangul = 10,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeHanziFull = 11,

		 /// <summary>
		 /// SupportByVersion MSForms 2
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("MSForms", 2)]
		 fmIMEModeHanzi = 12
	}
}