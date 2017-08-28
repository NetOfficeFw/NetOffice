using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsEnum)]
	public enum ParameterAttributesEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adParamSigned = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adParamNullable = 64,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adParamLong = 128
	}
}