using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// Interface IFont 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("BEF6E002-A874-101A-8BBA-00AA00300CAB")]
	public interface IFont : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		float Size { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Bold { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Italic { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Underline { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Strikethrough { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int16 Weight { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int16 Charset { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 hFont { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lplpfont">NetOffice.MSFormsApi.IFont lplpfont</param>
		[SupportByVersion("MSForms", 2)]
		Int32 Clone(out NetOffice.MSFormsApi.IFont lplpfont);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="lpFontOther">NetOffice.MSFormsApi.IFont lpFontOther</param>
		[SupportByVersion("MSForms", 2)]
		Int32 IsEqual(NetOffice.MSFormsApi.IFont lpFontOther);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="cyLogical">Int32 cyLogical</param>
		/// <param name="cyHimetric">Int32 cyHimetric</param>
		[SupportByVersion("MSForms", 2)]
		Int32 SetRatio(Int32 cyLogical, Int32 cyHimetric);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="hFont">Int32 hFont</param>
		[SupportByVersion("MSForms", 2)]
		Int32 AddRefHfont(Int32 hFont);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="hFont">Int32 hFont</param>
		[SupportByVersion("MSForms", 2)]
		Int32 ReleaseHfont(Int32 hFont);

		#endregion
	}
}
