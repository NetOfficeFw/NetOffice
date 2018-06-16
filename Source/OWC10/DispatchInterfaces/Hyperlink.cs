using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface Hyperlink 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("6F5A7620-0AC7-4DED-9A6B-A3547FD7B7BB")]
	public interface Hyperlink : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api.ISpreadsheet Application { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Address { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		NetOffice.OWC10Api._Range Parent { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string SubAddress { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void Delete();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		/// <param name="headerInfo">optional object headerInfo</param>
		[SupportByVersion("OWC10", 1)]
		void Follow(object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Follow();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Follow(object newWindow);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Follow(object newWindow, object addHistory);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Follow(object newWindow, object addHistory, object extraInfo);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		void Follow(object newWindow, object addHistory, object extraInfo, object method);

		#endregion
	}
}
