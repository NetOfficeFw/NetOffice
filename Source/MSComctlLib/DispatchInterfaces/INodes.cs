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
	/// DispatchInterface INodes 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "MSComctlLib", 6), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("C74190B7-8589-11D1-B16A-00C0F0283628")]
    [CoClassSource(typeof(NetOffice.MSComctlLibApi.Nodes))]
    public interface INodes : ICOMObject, IEnumerableProvider<NetOffice.MSComctlLibApi.INode>
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
		NetOffice.MSComctlLibApi.INode get_ControlDefault(object index);

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// Get/Set
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="value">NetOffice.MSComctlLibApi.INode value</param>
        [SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_ControlDefault(object index, NetOffice.MSComctlLibApi.INode value);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6), Redirect("get_ControlDefault")]
		NetOffice.MSComctlLibApi.INode ControlDefault(object index);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSComctlLibApi.INode this[object index] { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="relative">optional object relative</param>
		/// <param name="relationship">optional object relationship</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="image">optional object image</param>
		/// <param name="selectedImage">optional object selectedImage</param>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		NetOffice.MSComctlLibApi.INode Add(object relative, object relationship, object key, object text, object image, object selectedImage);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.INode Add();

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="relative">optional object relative</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.INode Add(object relative);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="relative">optional object relative</param>
		/// <param name="relationship">optional object relationship</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.INode Add(object relative, object relationship);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="relative">optional object relative</param>
		/// <param name="relationship">optional object relationship</param>
		/// <param name="key">optional object key</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.INode Add(object relative, object relationship, object key);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="relative">optional object relative</param>
		/// <param name="relationship">optional object relationship</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.INode Add(object relative, object relationship, object key, object text);

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="relative">optional object relative</param>
		/// <param name="relationship">optional object relationship</param>
		/// <param name="key">optional object key</param>
		/// <param name="text">optional object text</param>
		/// <param name="image">optional object image</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		NetOffice.MSComctlLibApi.INode Add(object relative, object relationship, object key, object text, object image);

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

        #region IEnumerable<NetOffice.MSComctlLibApi.INode>

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        new IEnumerator<NetOffice.MSComctlLibApi.INode> GetEnumerator();

        #endregion
    }
}
