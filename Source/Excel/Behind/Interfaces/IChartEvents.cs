using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IChartEvents 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IChartEvents : COMObject, NetOffice.ExcelApi.IChartEvents
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.ExcelApi.IChartEvents);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IChartEvents);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IChartEvents() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Activate()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Deactivate()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Deactivate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Resize()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Resize");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="button">Int32 button</param>
        /// <param name="shift">Int32 shift</param>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MouseDown(Int32 button, Int32 shift, Int32 x, Int32 y)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MouseDown", button, shift, x, y);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="button">Int32 button</param>
        /// <param name="shift">Int32 shift</param>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MouseUp(Int32 button, Int32 shift, Int32 x, Int32 y)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MouseUp", button, shift, x, y);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="button">Int32 button</param>
        /// <param name="shift">Int32 shift</param>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MouseMove(Int32 button, Int32 shift, Int32 x, Int32 y)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MouseMove", button, shift, x, y);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BeforeRightClick(bool cancel)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeRightClick", cancel);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DragPlot()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DragPlot");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DragOver()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DragOver");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        /// <param name="cancel">bool cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BeforeDoubleClick(Int32 elementID, Int32 arg1, Int32 arg2, bool cancel)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeDoubleClick", elementID, arg1, arg2, cancel);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Select(Int32 elementID, Int32 arg1, Int32 arg2)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Select", elementID, arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="seriesIndex">Int32 seriesIndex</param>
        /// <param name="pointIndex">Int32 pointIndex</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SeriesChange(Int32 seriesIndex, Int32 pointIndex)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SeriesChange", seriesIndex, pointIndex);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Calculate()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Calculate");
        }

        #endregion

        #pragma warning restore
    }
}
