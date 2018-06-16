using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSDATASRCApi
{
	/// <summary>
	/// Interface DataSource 
	/// SupportByVersion MSDATASRC, 4
	/// </summary>
	[SupportByVersion("MSDATASRC", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("7C0FFAB3-CD84-11D0-949A-00A0C91110ED")]
	public interface DataSource : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		/// <param name="riid">Guid riid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		object getDataMember(string bstrDM, Guid riid);

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		string getDataMemberName(Int32 lIndex);

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		Int32 getDataMemberCount();

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="pDSL">NetOffice.MSDATASRCApi.DataSourceListener pDSL</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		Int32 addDataSourceListener(NetOffice.MSDATASRCApi.DataSourceListener pDSL);

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="pDSL">NetOffice.MSDATASRCApi.DataSourceListener pDSL</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		Int32 removeDataSourceListener(NetOffice.MSDATASRCApi.DataSourceListener pDSL);

		#endregion
	}
}
