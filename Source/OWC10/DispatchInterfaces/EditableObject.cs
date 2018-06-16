using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface EditableObject 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("EB3286D3-226C-48F0-8049-2DB1E01DEE9C")]
	public interface EditableObject : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		object Value { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">object initialValue</param>
		/// <param name="arrowMode">bool arrowMode</param>
		/// <param name="caretPosition">Int32 caretPosition</param>
		[SupportByVersion("OWC10", 1)]
		void StartEdit(object initialValue, bool arrowMode, Int32 caretPosition);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="accept">bool accept</param>
		[SupportByVersion("OWC10", 1)]
		void EndEdit(bool accept);

		#endregion
	}
}
