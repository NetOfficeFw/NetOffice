using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface IDataAutoWrapper 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("EC72F590-F375-11CE-B9E8-00AA006B1A69")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.DataObject))]
    public interface IDataAutoWrapper : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void Clear();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="format">object format</param>
		[SupportByVersion("MSForms", 2)]
		bool GetFormat(object format);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="format">optional object format</param>
		[SupportByVersion("MSForms", 2)]
		string GetText(object format);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		string GetText();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("MSForms", 2)]
		void SetText(string text, object format);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		void SetText(string text);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void PutInClipboard();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		void GetFromClipboard();

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oKEffect">optional object oKEffect</param>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag(object oKEffect);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag();

		#endregion
	}
}
