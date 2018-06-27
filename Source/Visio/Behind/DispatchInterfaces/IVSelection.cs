using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVSelection 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	public class IVSelection : COMObject, NetOffice.VisioApi.IVSelection
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
                    _contractType = typeof(NetOffice.VisioApi.IVSelection);
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
                    _type = typeof(IVSelection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVSelection() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVShape get_Item16(Int16 index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Item16", typeof(NetOffice.VisioApi.IVShape), index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Item16
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Item16")]
		public virtual NetOffice.VisioApi.IVShape Item16(Int16 index)
		{
			return get_Item16(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 Count16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Count16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVPage ContainingPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ContainingPage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVMaster ContainingMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "ContainingMaster");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape ContainingShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "ContainingShape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Style");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string StyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string LineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string LineStyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineStyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FillStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FillStyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillStyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string TextStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string TextStyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextStyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 PersistsEvents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.VisioApi.IVShape this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 IterationMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "IterationMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IterationMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 get_ItemStatus(Int32 index)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ItemStatus", index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemStatus
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemStatus")]
		public virtual Int16 ItemStatus(Int32 index)
		{
			return get_ItemStatus(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape PrimaryItem
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "PrimaryItem");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public virtual stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
            }
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ContainingPageID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainingPageID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ContainingMasterID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainingMasterID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVMaster DataGraphic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "DataGraphic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataGraphic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVSelection SelectionForDragCopy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SelectionForDragCopy");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Export(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void BringForward()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BringForward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void BringToFront()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BringToFront");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SendBackward()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendBackward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SendToBack()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendToBack");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Combine()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Combine");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Fragment()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Fragment");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Intersect()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Intersect");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Subtract()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Subtract");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Union()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Union");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void FlipHorizontal()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FlipHorizontal");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void FlipVertical()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FlipVertical");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ReverseEnds()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReverseEnds");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate90()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate90");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void old_Copy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "old_Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void old_Cut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "old_Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void VoidDuplicate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "VoidDuplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void VoidGroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "VoidGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ConvertToGroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Ungroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Ungroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SelectAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void DeselectAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeselectAll");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="selectAction">Int16 selectAction</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select", sheetObject, selectAction);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Trim()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Trim");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Join()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Join");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void FitCurve(Double tolerance, Int16 flags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FitCurve", tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Layout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Layout");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true,true);
			lpr8Left = 0;
			lpr8Bottom = 0;
			lpr8Right = 0;
			lpr8Top = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top);
			Invoker.Method(this, "BoundingBox", paramsArray, modifiers);
			lpr8Left = (Double)paramsArray[1];
			lpr8Bottom = (Double)paramsArray[2];
			lpr8Right = (Double)paramsArray[3];
			lpr8Top = (Double)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		/// <param name="resultsMaster">optional object resultsMaster</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y, object resultsMaster)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", new object[]{ tolerance, flags, x, y, resultsMaster });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", tolerance, flags, x);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRegion", tolerance, flags, x, y);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape Group()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Group");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SwapEnds()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SwapEnds");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void AddToGroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void RemoveFromGroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveFromGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVSelection Duplicate()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "Duplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Copy(object flags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Copy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Cut(object flags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cut", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Cut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		/// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Move(Double dx, Double dy, object unitsNameOrCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", dx, dy, unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Move(Double dx, Double dy)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", dx, dy);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object pinUnitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY, object pinUnitsNameOrCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", new object[]{ angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX, pinY, pinUnitsNameOrCode });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", angle);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle, object angleUnitsNameOrCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", angle, angleUnitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", angle, angleUnitsNameOrCode, blastGuards);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", angle, angleUnitsNameOrCode, blastGuards, rotationType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", new object[]{ angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="angle">Double angle</param>
		/// <param name="angleUnitsNameOrCode">optional object angleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Rotate", new object[]{ angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX, pinY });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical, object glueToGuide)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Align", alignHorizontal, alignVertical, glueToGuide);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Align", alignHorizontal, alignVertical);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes distribute</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute, object glueToGuide)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Distribute", distribute, glueToGuide);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes distribute</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Distribute", distribute);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void UpdateAlignmentBox()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateAlignmentBox");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distance">Double distance</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Offset(Double distance)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Offset", distance);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ConnectShapes()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConnectShapes");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object pinUnitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY, object pinUnitsNameOrCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flip", new object[]{ flipDirection, flipType, blastGuards, pinX, pinY, pinUnitsNameOrCode });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flip", flipDirection);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flip", flipDirection, flipType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flip", flipDirection, flipType, blastGuards);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flip", flipDirection, flipType, blastGuards, pinX);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection flipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flip", new object[]{ flipDirection, flipType, blastGuards, pinX, pinY });
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void LinkToData(Int32 dataRecordsetID, Int32 dataRowID, object applyDataGraphicAfterLink)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkToData", dataRecordsetID, dataRowID, applyDataGraphicAfterLink);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void LinkToData(Int32 dataRecordsetID, Int32 dataRowID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkToData", dataRecordsetID, dataRowID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void BreakLinkToData(Int32 dataRecordsetID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BreakLinkToData", dataRecordsetID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual void GetIDs(out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)shapeIDs);
			Invoker.Method(this, "GetIDs", paramsArray, modifiers);
			shapeIDs = (Int32[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="columnNames">String[] columnNames</param>
		/// <param name="autoLinkFieldTypes">Int32[] autoLinkFieldTypes</param>
		/// <param name="fieldNames">String[] fieldNames</param>
		/// <param name="autoLinkBehavior">Int32 autoLinkBehavior</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32 AutomaticLink(Int32 dataRecordsetID, String[] columnNames, Int32[] autoLinkFieldTypes, String[] fieldNames, Int32 autoLinkBehavior, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)columnNames, (object)autoLinkFieldTypes, (object)fieldNames, autoLinkBehavior, (object)shapeIDs);
			object returnItem = Invoker.MethodReturn(this, "AutomaticLink", paramsArray);
			shapeIDs = (Int32[])paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="alignOrSpace">NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace</param>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical</param>
		/// <param name="spaceHorizontal">Double spaceHorizontal</param>
		/// <param name="spaceVertical">Double spaceVertical</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LayoutIncremental", new object[]{ alignOrSpace, alignHorizontal, alignVertical, spaceHorizontal, spaceVertical, unitCode });
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void LayoutChangeDirection(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LayoutChangeDirection", direction);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void AvoidPageBreaks()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AvoidPageBreaks");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection direction</param>
		/// <param name="distance">Double distance</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Resize", direction, distance, unitCode);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void AddToContainers()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToContainers");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void RemoveFromContainers()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveFromContainers");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop, newShape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="delFlags">Int32 delFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void DeleteEx(Int32 delFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteEx", delFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] GetContainers(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetContainers", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] GetCallouts(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetCallouts", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] MemberOfContainersUnion()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "MemberOfContainersUnion", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] MemberOfContainersIntersection()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "MemberOfContainersIntersection", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType formatType</param>
		/// <param name="formatValue">optional object FormatValue = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType, object formatValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatType, formatValue);
			object returnItem = (object)Invoker.MethodReturn(this, "SetContainerFormat", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType formatType</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatType);
			object returnItem = (object)Invoker.MethodReturn(this, "SetContainerFormat", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		/// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
		[SupportByVersion("Visio", 15, 16)]
		public virtual NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop, replaceFlags);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public virtual NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
		/// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
		/// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
		/// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
		/// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
		/// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
		/// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
		/// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
		[SupportByVersion("Visio", 15, 16)]
		public virtual void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetQuickStyle", new object[]{ lineMatrix, fillMatrix, effectsMatrix, fontMatrix, lineColor, fillColor, shadowColor, fontColor });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVShape>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVShape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVShape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVShape>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.VisioApi.IVShape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVShape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

