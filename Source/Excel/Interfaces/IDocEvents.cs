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
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IDocEvents : COMObject
	{
		#pragma warning disable

		#region Type Information

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
                    _type = typeof(IDocEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDocEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SelectionChange(NetOffice.ExcelApi.Range target)
		{
			return Factory.ExecuteInt32MethodGet(this, "SelectionChange", target);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeDoubleClick(NetOffice.ExcelApi.Range target, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "BeforeDoubleClick", target, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 BeforeRightClick(NetOffice.ExcelApi.Range target, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "BeforeRightClick", target, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Activate()
		{
			return Factory.ExecuteInt32MethodGet(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Deactivate()
		{
			return Factory.ExecuteInt32MethodGet(this, "Deactivate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Calculate()
		{
			return Factory.ExecuteInt32MethodGet(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Range target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Change(NetOffice.ExcelApi.Range target)
		{
			return Factory.ExecuteInt32MethodGet(this, "Change", target);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.Hyperlink target</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 FollowHyperlink(NetOffice.ExcelApi.Hyperlink target)
		{
			return Factory.ExecuteInt32MethodGet(this, "FollowHyperlink", target);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 PivotTableUpdate(NetOffice.ExcelApi.PivotTable target)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableUpdate", target);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="targetRange">NetOffice.ExcelApi.Range targetRange</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 PivotTableAfterValueChange(NetOffice.ExcelApi.PivotTable targetPivotTable, NetOffice.ExcelApi.Range targetRange)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableAfterValueChange", targetPivotTable, targetRange);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 PivotTableBeforeAllocateChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableBeforeAllocateChanges", targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 PivotTableBeforeCommitChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, bool cancel)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableBeforeCommitChanges", targetPivotTable, valueChangeStart, valueChangeEnd, cancel);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="targetPivotTable">NetOffice.ExcelApi.PivotTable targetPivotTable</param>
		/// <param name="valueChangeStart">Int32 valueChangeStart</param>
		/// <param name="valueChangeEnd">Int32 valueChangeEnd</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 PivotTableBeforeDiscardChanges(NetOffice.ExcelApi.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableBeforeDiscardChanges", targetPivotTable, valueChangeStart, valueChangeEnd);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.PivotTable target</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 PivotTableChangeSync(NetOffice.ExcelApi.PivotTable target)
		{
			return Factory.ExecuteInt32MethodGet(this, "PivotTableChangeSync", target);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 LensGalleryRenderComplete()
		{
			return Factory.ExecuteInt32MethodGet(this, "LensGalleryRenderComplete");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="target">NetOffice.ExcelApi.TableObject target</param>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 TableUpdate(NetOffice.ExcelApi.TableObject target)
		{
			return Factory.ExecuteInt32MethodGet(this, "TableUpdate", target);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		[SupportByVersion("Excel", 15, 16)]
		public Int32 BeforeDelete()
		{
			return Factory.ExecuteInt32MethodGet(this, "BeforeDelete");
		}

		#endregion

		#pragma warning restore
	}
}
