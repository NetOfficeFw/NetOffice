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
	/// Default implementation of <see cref="NetOffice.VisioApi.EventContracts.EMasters"/>
	/// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class EMasters_SinkHelper : SinkHelper, NetOffice.VisioApi.EventContracts.EMasters
	{
		#region Static
		
		/// <summary>
		/// Interface Id from EMasters
		/// </summary>
		public static readonly string Id = "000D0B07-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        /// <exception cref="NetOfficeCOMException">Unexpected error</exception>
        public EMasters_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region EMasters

        /// <summary>
        /// 
        /// </summary>
        /// <param name="master"></param>
        public void MasterAdded([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("MasterAdded"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("MasterAdded", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void MasterChanged([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("MasterChanged"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("MasterChanged", ref paramsArray);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void BeforeMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
		{
            if (!Validate("BeforeMasterDelete"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
			paramsArray[0] = newMaster;
			EventBinding.RaiseCustomEvent("BeforeMasterDelete", ref paramsArray);
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
		/// <param name="master"></param>
        public void QueryCancelMasterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("QueryCancelMasterDelete"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("QueryCancelMasterDelete", ref paramsArray);
        }

		/// <summary>
		/// 
		/// </summary>
		/// <param name="master"></param>
        public void MasterDeleteCanceled([In, MarshalAs(UnmanagedType.IDispatch)] object master)
        {
            if (!Validate("MasterDeleteCanceled"))
            {
                Invoker.ReleaseParamsArray(master);
                return;
            }

            NetOffice.VisioApi.IVMaster newMaster = Factory.CreateEventArgumentObjectFromComProxy(EventClass, master) as NetOffice.VisioApi.IVMaster;
            object[] paramsArray = new object[1];
            paramsArray[0] = newMaster;
            EventBinding.RaiseCustomEvent("MasterDeleteCanceled", ref paramsArray);
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

        #endregion
    }
	
}
