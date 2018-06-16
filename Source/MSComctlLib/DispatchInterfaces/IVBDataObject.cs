using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IVBDataObject 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("2334D2B1-713E-11CF-8AE5-00AA00C00905")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.DataObject))]
    public interface IVBDataObject : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.IVBDataObjectFiles Files { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sFormat">Int16 sFormat</param>
		[SupportByVersion("MSComctlLib", 6)]
		object GetData(Int16 sFormat);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sFormat">Int16 sFormat</param>
		[SupportByVersion("MSComctlLib", 6)]
		bool GetFormat(Int16 sFormat);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="vValue">optional object vValue</param>
		/// <param name="vFormat">optional object vFormat</param>
		[SupportByVersion("MSComctlLib", 6)]
		void SetData(object vValue, object vFormat);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		void SetData();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="vValue">optional object vValue</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		void SetData(object vValue);

		#endregion
	}
}
