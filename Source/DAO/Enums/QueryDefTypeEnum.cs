using System;
using NetOffice;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 12, 3.6
	 /// </summary>
	[SupportByVersionAttribute("DAO", 12,3.6)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum QueryDefTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQSelect = 0,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>224</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQProcedure = 224,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQAction = 240,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQCrosstab = 16,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQDelete = 32,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQUpdate = 48,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQAppend = 64,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQMakeTable = 80,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQDDL = 96,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQSQLPassThrough = 112,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQSetOperation = 128,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQSPTBulk = 144,

		 /// <summary>
		 /// SupportByVersion DAO 12, 3.6
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersionAttribute("DAO", 12,3.6)]
		 dbQCompound = 160
	}
}