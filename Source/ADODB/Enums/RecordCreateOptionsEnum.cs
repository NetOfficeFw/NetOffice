using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.5
	 /// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsEnum)]
	public enum RecordCreateOptionsEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adCreateCollection = 8192,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>-2147483648</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adCreateStructDoc = -2147483648,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adCreateNonCollection = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>33554432</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adOpenIfExists = 33554432,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>67108864</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adCreateOverwrite = 67108864,

		 /// <summary>
		 /// SupportByVersion ADODB 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("ADODB", 2.5)]
		 adFailIfNotExists = -1
	}
}