using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPStrings 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("914934A9-5A91-11CF-8700-00AA0060263B")]
	public interface PPStrings : Collection
	{
		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		string this[Int32 index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="_string">string string</param>
		[SupportByVersion("PowerPoint", 9)]
		string Add(string _string);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="_string">string string</param>
		/// <param name="position">Int32 position</param>
		[SupportByVersion("PowerPoint", 9)]
		void Insert(string _string, Int32 position);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9)]
		void Delete(Int32 index);

		#endregion
	}
}
