using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface Tabs 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSForms", 2), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("944ACF93-A1E6-11CE-8104-00AA00611080")]
	public interface Tabs : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab _GetItemByIndex(Int32 lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab _GetItemByName(string bstr);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		object this[object varg] { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		object Enum();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="lIndex">optional object lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab Add(object bstrName, object bstrCaption, object lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab Add();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab Add(object bstrName);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab Add(object bstrName, object bstrCaption);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab _Add(string bstrName, string bstrCaption);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tab _Insert(string bstrName, string bstrCaption, Int32 lIndex);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersion("MSForms", 2)]
		void Remove(object varg);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Clear();

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion MSForms, 2
        /// </summary>
        [SupportByVersion("MSForms", 2)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
