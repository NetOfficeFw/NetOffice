using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface DesignAdviseSink 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("9B3E2331-87A6-11D1-BACD-00C04FAC6863")]
	public interface DesignAdviseSink : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="fGrid">Int32 fGrid</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ObjectAdded(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, Int32 fGrid);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ObjectDeleted(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrRsd">string bstrRsd</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ObjectMoved(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrRsd);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 DataModelLoad();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ObjectChanged(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ObjectDeleteComplete(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrPreviousName">string bstrPreviousName</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ObjectRenamed(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrPreviousName);

		#endregion
	}
}
