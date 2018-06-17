using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// DispatchInterface ITabStrip 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("04598FC2-866C-11CF-AB7C-00AA00C08FCF")]
    [CoClassSource(typeof(NetOffice.MSFormsApi.TabStrip))]
    public interface ITabStrip : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 BackColor { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 ForeColor { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Font _Font_Reserved { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		NetOffice.MSFormsApi.Font Font { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string FontName { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool FontBold { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool FontItalic { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool FontUnderline { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		bool FontStrikethru { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		float FontSize { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool Enabled { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2), NativeResult]
		stdole.Picture MouseIcon { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmMousePointer MousePointer { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		bool MultiRow { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmTabStyle Style { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Enums.fmTabOrientation TabOrientation { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ClientTop { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ClientLeft { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ClientWidth { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single ClientHeight { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		NetOffice.MSFormsApi.Tabs Tabs { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSFormsApi.Tab SelectedItem { get; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Int32 Value { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single TabFixedWidth { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		Single TabFixedHeight { get; set; }

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 FontWeight { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedWidth">Int32 tabFixedWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _SetTabFixedWidth(Int32 tabFixedWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedWidth">Int32 tabFixedWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _GetTabFixedWidth(out Int32 tabFixedWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedHeight">Int32 tabFixedHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _SetTabFixedHeight(Int32 tabFixedHeight);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="tabFixedHeight">Int32 tabFixedHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _GetTabFixedHeight(out Int32 tabFixedHeight);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientTop">Int32 clientTop</param>
		[SupportByVersion("MSForms", 2)]
		void _GetClientTop(out Int32 clientTop);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientLeft">Int32 clientLeft</param>
		[SupportByVersion("MSForms", 2)]
		void _GetClientLeft(out Int32 clientLeft);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientWidth">Int32 clientWidth</param>
		[SupportByVersion("MSForms", 2)]
		void _GetClientWidth(out Int32 clientWidth);

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="clientHeight">Int32 clientHeight</param>
		[SupportByVersion("MSForms", 2)]
		void _GetClientHeight(out Int32 clientHeight);

		#endregion
	}
}
