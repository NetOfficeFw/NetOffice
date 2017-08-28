using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.DAOApi.Enums
{
	 /// <summary>
	 /// SupportByVersion DAO 3.6, 12.0
	 /// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsEnum)]
	public enum QueryDefTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQSelect = 0,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>224</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQProcedure = 224,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>240</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQAction = 240,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQCrosstab = 16,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQDelete = 32,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQUpdate = 48,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQAppend = 64,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQMakeTable = 80,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQDDL = 96,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQSQLPassThrough = 112,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQSetOperation = 128,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQSPTBulk = 144,

		 /// <summary>
		 /// SupportByVersion DAO 3.6, 12.0
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersion("DAO", 3.6,12.0)]
		 dbQCompound = 160
	}
}