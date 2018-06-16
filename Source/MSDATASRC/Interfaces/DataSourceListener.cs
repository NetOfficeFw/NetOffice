using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSDATASRCApi
{
	/// <summary>
	/// Interface DataSourceListener 
	/// SupportByVersion MSDATASRC, 4
	/// </summary>
	[SupportByVersion("MSDATASRC", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("7C0FFAB2-CD84-11D0-949A-00A0C91110ED")]
	public interface DataSourceListener : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		Int32 dataMemberChanged(string bstrDM);

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		Int32 dataMemberAdded(string bstrDM);

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		Int32 dataMemberRemoved(string bstrDM);

		#endregion
	}
}
