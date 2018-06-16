using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Recordset21_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00000555-0000-0010-8000-00AA006D2EA4")]
	public interface Recordset21_Deprecated : Recordset20_Deprecated
	{
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
        new NetOffice.ADODBApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new object get_Collect(object index);

        /// <summary>
        /// SupportByVersion ADODB 2.5
        /// Get/Set
        /// </summary>
        /// <param name="index">object index</param>
        /// <param name="value">object value</param>
        [SupportByVersion("ADODB", 2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        new void set_Collect(object index, object value);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Alias for get_Collect
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5), Redirect("get_Collect")]
        new object Collect(object index);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		string Index { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		/// <param name="seekOption">optional NetOffice.ADODBApi.Enums.SeekEnum SeekOption = 1</param>
		[SupportByVersion("ADODB", 2.5)]
		void Seek(object keyValues, object seekOption);

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		void Seek(object keyValues);

		#endregion
	}
}
