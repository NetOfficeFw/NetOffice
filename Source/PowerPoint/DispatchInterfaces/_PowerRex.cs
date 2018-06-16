using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface _PowerRex 
	/// SupportByVersion PowerPoint, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("914934D3-5A91-11CF-8700-00AA0060263B")]
    [CoClassSource(typeof(NetOffice.PowerPointApi.PowerRex))]
    public interface _PowerRex : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="erorCode">object erorCode</param>
		/// <param name="bstrErrorDesc">object bstrErrorDesc</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void OnAsfEncoderEvent(object erorCode, object bstrErrorDesc);

		#endregion
	}
}
