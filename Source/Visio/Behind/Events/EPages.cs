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
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EPages"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EPages_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EPages
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EPages
		/// </summary>
		public static readonly string Id = "000D0B09-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EPages_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EPages

        /// <summary>
        /// 
        /// </summary>
        /// <param name="page"></param>
        public void PageAdded([In, MarshalAs(UnmanagedType.IDispatch)] object page)
		{
            if (!Validate("PageAdded"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
			paramsArray[0] = newPage;
			EventBinding.RaiseCustomEvent("PageAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void PageChanged([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("PageChanged"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("PageChanged", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void BeforePageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("BeforePageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("BeforePageDelete", ref paramsArray);
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
		/// <param name="connects"></param>
        public void ConnectionsAdded([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
        {
            if (!Validate("ConnectionsAdded"))
            {
                Invoker.ReleaseParamsArray(connects);
                return;
            }

            NetOffice.VisioApi.IVConnects newConnects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, connects) as NetOffice.VisioApi.IVConnects;
            object[] paramsArray = new object[1];
            paramsArray[0] = newConnects;
            EventBinding.RaiseCustomEvent("ConnectionsAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="connects"></param>
        public void ConnectionsDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object connects)
        {
            if (!Validate("ConnectionsDeleted"))
            {
                Invoker.ReleaseParamsArray(connects);
                return;
            }

            NetOffice.VisioApi.IVConnects newConnects = Factory.CreateEventArgumentObjectFromComProxy(EventClass, connects) as NetOffice.VisioApi.IVConnects;
            object[] paramsArray = new object[1];
            paramsArray[0] = newConnects;
            EventBinding.RaiseCustomEvent("ConnectionsDeleted", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void QueryCancelPageDelete([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("QueryCancelPageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("QueryCancelPageDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="page"></param>
        public void PageDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object page)
        {
            if (!Validate("QueryCancelPageDelete"))
            {
                Invoker.ReleaseParamsArray(page);
                return;
            }

            NetOffice.VisioApi.IVPage newPage = Factory.CreateEventArgumentObjectFromComProxy(EventClass, page) as NetOffice.VisioApi.IVPage;
            object[] paramsArray = new object[1];
            paramsArray[0] = newPage;
            EventBinding.RaiseCustomEvent("PageDeleteCanceled", ref paramsArray);
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

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void ContainerRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
        {
            if (!Validate("ContainerRelationshipAdded"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShapePair;
            EventBinding.RaiseCustomEvent("ContainerRelationshipAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void ContainerRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
        {
            if (!Validate("ContainerRelationshipDeleted"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShapePair;
            EventBinding.RaiseCustomEvent("ContainerRelationshipDeleted", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void CalloutRelationshipAdded([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
        {
            if (!Validate("CalloutRelationshipAdded"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShapePair;
            EventBinding.RaiseCustomEvent("CalloutRelationshipAdded", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="shapePair"></param>
        public void CalloutRelationshipDeleted([In, MarshalAs(UnmanagedType.IDispatch)] object shapePair)
        {
            if (!Validate("CalloutRelationshipDeleted"))
            {
                Invoker.ReleaseParamsArray(shapePair);
                return;
            }

            NetOffice.VisioApi.IVRelatedShapePairEvent newShapePair = Factory.CreateEventArgumentObjectFromComProxy(EventClass, shapePair) as NetOffice.VisioApi.IVRelatedShapePairEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newShapePair;
            EventBinding.RaiseCustomEvent("CalloutRelationshipDeleted", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="replaceShapes"></param>
        public void QueryCancelReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
        {
            if (!Validate("QueryCancelReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newreplaceShapes;
            EventBinding.RaiseCustomEvent("QueryCancelReplaceShapes", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="replaceShapes"></param>
        public void ReplaceShapesCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
        {
            if (!Validate("QueryCancelReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newreplaceShapes;
            EventBinding.RaiseCustomEvent("ReplaceShapesCanceled", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="replaceShapes"></param>
        public void BeforeReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object replaceShapes)
        {
            if (!Validate("BeforeReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(replaceShapes);
                return;
            }

            NetOffice.VisioApi.IVReplaceShapesEvent newreplaceShapes = Factory.CreateEventArgumentObjectFromComProxy(EventClass, replaceShapes) as NetOffice.VisioApi.IVReplaceShapesEvent;
            object[] paramsArray = new object[1];
            paramsArray[0] = newreplaceShapes;
            EventBinding.RaiseCustomEvent("BeforeReplaceShapes", ref paramsArray);
        }
        
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sel"></param>
        public void AfterReplaceShapes([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
        {
            if (!Validate("AfterReplaceShapes"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

            NetOffice.VisioApi.IVSelection newsel = Factory.CreateEventArgumentObjectFromComProxy(EventClass, sel) as NetOffice.VisioApi.IVSelection;
            object[] paramsArray = new object[1];
            paramsArray[0] = newsel;
            EventBinding.RaiseCustomEvent("AfterReplaceShapes", ref paramsArray);
        }

        #endregion
    }

}
