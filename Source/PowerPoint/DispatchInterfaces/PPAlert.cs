using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPAlert 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("9149349F-5A91-11CF-8700-00AA0060263B")]
	public interface PPAlert : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("PowerPoint", 9), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		Int32 PressedButton { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		string OnButton { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="title">string title</param>
		/// <param name="type">Int32 type</param>
		/// <param name="text">string text</param>
		/// <param name="leftBtn">string leftBtn</param>
		/// <param name="middleBtn">string middleBtn</param>
		/// <param name="rightBtn">string rightBtn</param>
		[SupportByVersion("PowerPoint", 9)]
		void Run(string title, Int32 type, string text, string leftBtn, string middleBtn, string rightBtn);

		#endregion
	}
}
