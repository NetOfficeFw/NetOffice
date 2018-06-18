using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Exceptions;

namespace NetOffice.VisioApi.Behind.EventContracts
{

	/// <summary>
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EShape"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EShape_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EShape
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EShape
		/// </summary>
		public static readonly string Id = "000D0B0B-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EShape_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EShape

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        public void CellChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
		{
            if (!Validate("CellChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
			paramsArray[0] = newCell;
			EventBinding.RaiseCustomEvent("CellChanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeAdded"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void BeforeSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("BeforeSelectionDelete"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("BeforeSelectionDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void SelectionAdded([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("SelectionAdded"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("SelectionAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void BeforeShapeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("BeforeShapeDelete"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("BeforeShapeDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void TextChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("TextChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("TextChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="cell"></param>
        public void FormulaChanged([In, MarshalAs(UnmanagedType.IDispatch)] object cell)
        {
            if (!Validate("FormulaChanged"))
            {
                Invoker.ReleaseParamsArray(cell);
                return;
            }

            NetOffice.VisioApi.IVCell newCell = Factory.CreateEventArgumentObjectFromComProxy(EventClass, cell) as NetOffice.VisioApi.IVCell;
            object[] paramsArray = new object[1];
            paramsArray[0] = newCell;
            EventBinding.RaiseCustomEvent("FormulaChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeParentChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeParentChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeParentChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void BeforeShapeTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("BeforeShapeTextEdit"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("BeforeShapeTextEdit", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeExitedTextEdit([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeExitedTextEdit"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeExitedTextEdit", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelSelectionDelete([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelSelectionDelete"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelSelectionDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void SelectionDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("SelectionDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("SelectionDeleteCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelUngroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelUngroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelUngroup", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void UngroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelUngroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("UngroupCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelConvertToGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelConvertToGroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelConvertToGroup", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void ConvertToGroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("ConvertToGroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("ConvertToGroupCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void QueryCancelGroup([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("QueryCancelGroup"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("QueryCancelGroup", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="selection"></param>
        public void GroupCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object selection)
        {
            if (!Validate("GroupCanceled"))
            {
                Invoker.ReleaseParamsArray(selection);
                return;
            }

            NetOffice.VisioApi.IVSelection newSelection = Factory.CreateEventArgumentObjectFromComProxy(EventClass, selection) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSelection;
            EventBinding.RaiseCustomEvent("GroupCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
        public void ShapeDataGraphicChanged([In, MarshalAs(UnmanagedType.IDispatch)] object shape)
        {
            if (!Validate("ShapeDataGraphicChanged"))
            {
                Invoker.ReleaseParamsArray(shape);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShape;
            EventBinding.RaiseCustomEvent("ShapeDataGraphicChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
		/// <param name="dataRecordsetID"></param>
		/// <param name="dataRowID"></param>
        public void ShapeLinkAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
        {
            if (!Validate("ShapeLinkAdded"))
            {
                Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            Int32 newDataRecordsetID = ToInt32(dataRecordsetID);
			Int32 newDataRowID = ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			EventBinding.RaiseCustomEvent("ShapeLinkAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shape"></param>
		/// <param name="dataRecordsetID"></param>
		/// <param name="dataRowID"></param>
        public void ShapeLinkDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shape, [In] object dataRecordsetID, [In] object dataRowID)
		{
            if (!Validate("ShapeLinkDeleted"))
            {
                Invoker.ReleaseParamsArray(shape, dataRecordsetID, dataRowID);
                return;
            }

            NetOffice.VisioApi.IVShape newShape = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shape) as NetOffice.VisioApi.IVShape;
            Int32 newDataRecordsetID = ToInt32(dataRecordsetID);
			Int32 newDataRowID = ToInt32(dataRowID);
			object[] paramsArray = new object[3];
			paramsArray[0] = newShape;
			paramsArray[1] = newDataRecordsetID;
			paramsArray[2] = newDataRowID;
			EventBinding.RaiseCustomEvent("ShapeLinkDeleted", ref paramsArray);
		}

		#endregion
	}
	
}
