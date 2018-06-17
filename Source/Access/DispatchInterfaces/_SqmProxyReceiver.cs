using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _SqmProxyReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("F7CF612C-79E8-46EE-AE58-E589E5B7D6A0")]
	public interface _SqmProxyReceiver : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void SetDataPoint(UIntPtr id, UIntPtr dwValue);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void SetDataPointMax(UIntPtr id, UIntPtr dwValue);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void SetDataPointMin(UIntPtr id, UIntPtr dwValue);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="type">UIntPtr type</param>
		/// <param name="width">UIntPtr width</param>
		/// <param name="maxRows">UIntPtr maxRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void CreateStream(UIntPtr id, UIntPtr type, UIntPtr width, UIntPtr maxRows);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void AddStreamData1(UIntPtr id, UIntPtr dw1);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		/// <param name="dw2">UIntPtr dw2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void AddStreamData2(UIntPtr id, UIntPtr dw1, UIntPtr dw2);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		/// <param name="dw2">UIntPtr dw2</param>
		/// <param name="dw3">UIntPtr dw3</param>
		/// <param name="dw4">UIntPtr dw4</param>
		/// <param name="dw5">UIntPtr dw5</param>
		/// <param name="dw6">UIntPtr dw6</param>
		/// <param name="dw7">UIntPtr dw7</param>
		/// <param name="dw8">UIntPtr dw8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void AddStreamData8(UIntPtr id, UIntPtr dw1, UIntPtr dw2, UIntPtr dw3, UIntPtr dw4, UIntPtr dw5, UIntPtr dw6, UIntPtr dw7, UIntPtr dw8);

		#endregion
	}
}
