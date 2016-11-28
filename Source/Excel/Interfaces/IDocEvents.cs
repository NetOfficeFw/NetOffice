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
	/// Interface IDocEvents 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IDocEvents : COMObject
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
                    _type = typeof(IDocEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDocEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocEvents(string progId) : base(progId)
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
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 SelectionChange(NetOffice.ExcelApi.Range target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "SelectionChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeDoubleClick(NetOffice.ExcelApi.Range target, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeDoubleClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeRightClick(NetOffice.ExcelApi.Range target, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeRightClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

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
		public Int32 Calculate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Calculate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range Target</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Change(NetOffice.ExcelApi.Range target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "Change", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Hyperlink Target</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 FollowHyperlink(NetOffice.ExcelApi.Hyperlink target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "FollowHyperlink", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public Int32 PivotTableUpdate(NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "PivotTableUpdate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="targetRange">NetOffice.ExcelApi.Range TargetRange</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 PivotTableAfterValueChange(NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetPivotTable, targetRange);
			object returnItem = Invoker.MethodReturn(this, "PivotTableAfterValueChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="valueChangeStart">Int32 ValueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 ValueChangeEnd</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 PivotTableBeforeAllocateChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
			object returnItem = Invoker.MethodReturn(this, "PivotTableBeforeAllocateChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="valueChangeStart">Int32 ValueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 ValueChangeEnd</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 PivotTableBeforeCommitChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
			object returnItem = Invoker.MethodReturn(this, "PivotTableBeforeCommitChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable TargetPivotTable</param>
		/// <param name="valueChangeStart">Int32 ValueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 ValueChangeEnd</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 PivotTableBeforeDiscardChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(targetPivotTable, valueChangeStart, valueChangeEnd);
			object returnItem = Invoker.MethodReturn(this, "PivotTableBeforeDiscardChanges", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable Target</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public Int32 PivotTableChangeSync(NetOffice.ExcelApi.PivotTable target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "PivotTableChangeSync", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 LensGalleryRenderComplete()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "LensGalleryRenderComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.TableObject Target</param>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 TableUpdate(NetOffice.ExcelApi.TableObject target)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target);
			object returnItem = Invoker.MethodReturn(this, "TableUpdate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public Int32 BeforeDelete()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BeforeDelete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}