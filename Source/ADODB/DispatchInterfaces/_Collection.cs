using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Collection 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "ADODB", 2.1, 2.5)]
	[TypeId("00000512-0000-0010-8000-00AA006D2EA4")]
	public interface _Collection : ICOMObject, IEnumerableProvider<object>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Refresh();

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        new IEnumerator<object> GetEnumerator();

        #endregion
    }
}
