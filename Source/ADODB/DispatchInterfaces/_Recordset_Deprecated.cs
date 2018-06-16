using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Recordset_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000556-0000-0010-8000-00AA006D2EA4")]
	public interface _Recordset_Deprecated : Recordset21_Deprecated
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new NetOffice.ADODBApi.Properties Properties { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destination">optional object destination</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[SupportByVersion("ADODB", 2.5)]
		void Save(object destination, object persistFormat);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Save();

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Save(object destination);

		#endregion
	}
}
