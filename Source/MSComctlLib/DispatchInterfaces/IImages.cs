using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IImages 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("2C247F24-8591-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.ListImages))]
    public interface IImages : ICOMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.IImage>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		Int16 Count { get; set; }

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSComctlLibApi.IImage get_ControlDefault(object index);

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// Get/Set
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="value">NetOffice.MSComctlLibApi.IImage value</param>
        [SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_ControlDefault(object index, NetOffice.MSComctlLibApi.IImage value);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		NetOffice.MSComctlLibApi.IImage ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSComctlLibApi.IImage this[object index] { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		/// <param name="picture">optional object picture</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.IImage Add(object index, object key, object picture);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IImage Add();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IImage Add(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.IImage Add(object index, object key);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		void Remove(object index);

        #endregion

        #region IEnumerable<NetOffice.MSComctlLibApi.IImage>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<NetOffice.MSComctlLibApi.IImage> GetEnumerator();

        #endregion
    }
}
