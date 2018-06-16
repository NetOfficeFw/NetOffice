using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface _DynaCollection 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000000A2-0000-0010-8000-00AA006D2EA4")]
	public interface _DynaCollection : _Collection
	{
		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="_object">object object</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Append(object _object);

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		void Delete(string name);

		#endregion
	}
}
