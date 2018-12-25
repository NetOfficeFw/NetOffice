using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj997235.aspx </remarks>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoPictureCompress
	{
		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("Office", 16)]
		 msoPictureCompressDocDefault = -1,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 16)]
		 msoPictureCompressFalse = 0,

		 /// <summary>
		 /// SupportByVersion Office 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 16)]
		 msoPictureCompressTrue = 1
	}
}