using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Collection
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
	[TypeId("91493450-5A91-11CF-8700-00AA0060263B")]
	public interface Collection : ICOMObject, IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		object _Index(Int32 index);

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
