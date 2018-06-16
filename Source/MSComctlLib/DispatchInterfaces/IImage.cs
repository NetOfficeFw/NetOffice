using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IImage 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("2C247F26-8591-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.ListImage))]
    public interface IImage : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		Int16 Index { get; set; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		string Key { get; set; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		object Tag { get; set; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		stdole.Picture Picture { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		/// <param name="style">optional object style</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Draw(Int32 hDC, object x, object y, object style);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		void Draw(Int32 hDC);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="x">optional object x</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		void Draw(Int32 hDC, object x);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		void Draw(Int32 hDC, object x, object y);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		stdole.Picture ExtractIcon();

		#endregion
	}
}
