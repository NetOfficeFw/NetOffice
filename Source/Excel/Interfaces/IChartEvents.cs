using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// Interface IChartEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IChartEvents : COMObject
	{
		#pragma warning disable
		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IChartEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IChartEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IChartEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IChartEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IChartEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IChartEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IChartEvents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Activate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Activate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Deactivate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Deactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Resize()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Resize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="button">Int32 Button</param>
		/// <param name="shift">Int32 Shift</param>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MouseDown(Int32 button, Int32 shift, Int32 x, Int32 y)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(button, shift, x, y);
			object returnItem = Invoker.MethodReturn(this, "MouseDown", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="button">Int32 Button</param>
		/// <param name="shift">Int32 Shift</param>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MouseUp(Int32 button, Int32 shift, Int32 x, Int32 y)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(button, shift, x, y);
			object returnItem = Invoker.MethodReturn(this, "MouseUp", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="button">Int32 Button</param>
		/// <param name="shift">Int32 Shift</param>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 MouseMove(Int32 button, Int32 shift, Int32 x, Int32 y)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(button, shift, x, y);
			object returnItem = Invoker.MethodReturn(this, "MouseMove", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeRightClick(bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeRightClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 DragPlot()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DragPlot", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 DragOver()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DragOver", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="elementID">Int32 ElementID</param>
		/// <param name="arg1">Int32 Arg1</param>
		/// <param name="arg2">Int32 Arg2</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeDoubleClick(Int32 elementID, Int32 arg1, Int32 arg2, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(elementID, arg1, arg2, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeDoubleClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="elementID">Int32 ElementID</param>
		/// <param name="arg1">Int32 Arg1</param>
		/// <param name="arg2">Int32 Arg2</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Select(Int32 elementID, Int32 arg1, Int32 arg2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(elementID, arg1, arg2);
			object returnItem = Invoker.MethodReturn(this, "Select", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="seriesIndex">Int32 SeriesIndex</param>
		/// <param name="pointIndex">Int32 PointIndex</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SeriesChange(Int32 seriesIndex, Int32 pointIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(seriesIndex, pointIndex);
			object returnItem = Invoker.MethodReturn(this, "SeriesChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Calculate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Calculate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}