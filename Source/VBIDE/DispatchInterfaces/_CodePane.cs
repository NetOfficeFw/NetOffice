using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
	/// DispatchInterface _CodePane
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0002E176-0000-0000-C000-000000000046")]
    public interface _CodePane : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.CodePanes Collection { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.VBE VBE { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Window Window { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 TopLine { get; set; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        Int32 CountOfVisibleLines { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.CodeModule CodeModule { get; }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        NetOffice.VBIDEApi.Enums.vbext_CodePaneview CodePaneView { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void GetSelection(out Int32 startLine, out Int32 startColumn, out Int32 endLine, out Int32 endColumn);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void SetSelection(Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn);

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        void Show();

        #endregion
    }
}
