using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IDocEvents 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IDocEvents : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 SelectionChange(NetOffice.ExcelApi.Range target);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BeforeDoubleClick(NetOffice.ExcelApi.Range target, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BeforeRightClick(NetOffice.ExcelApi.Range target, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Activate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Deactivate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Calculate();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.Range target</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Change(NetOffice.ExcelApi.Range target);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.Hyperlink target</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 FollowHyperlink(NetOffice.ExcelApi.Hyperlink target);

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        Int32 PivotTableUpdate(NetOffice.ExcelApi.PivotTable target);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="targetRange">NetOffice.ExcelApi.Range targetRange</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 PivotTableAfterValueChange(NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="valueChangeStart">Int32 valueChangeStart</param>
        /// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 PivotTableBeforeAllocateChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="valueChangeStart">Int32 valueChangeStart</param>
        /// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 PivotTableBeforeCommitChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel);


        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
        /// <param name="valueChangeStart">Int32 valueChangeStart</param>
        /// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 PivotTableBeforeDiscardChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        Int32 PivotTableChangeSync(NetOffice.ExcelApi.PivotTable target);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 LensGalleryRenderComplete();

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="target">NetOffice.ExcelApi.TableObject target</param>
        [SupportByVersion("Excel", 15, 16)]
        Int32 TableUpdate(NetOffice.ExcelApi.TableObject target);

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [SupportByVersion("Excel", 15, 16)]
        Int32 BeforeDelete();

        #endregion
    }
}
