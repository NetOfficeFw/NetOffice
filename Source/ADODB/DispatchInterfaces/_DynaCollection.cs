using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _DynaCollection 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000513-0000-0010-8000-00AA006D2EA4")]
	public interface _DynaCollection : _Collection
	{
		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="_object">object object</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Append(object _object);

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		void Delete(object index);

		#endregion
	}
}
