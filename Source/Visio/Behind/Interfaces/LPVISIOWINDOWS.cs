using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOWINDOWS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	public class LPVISIOWINDOWS : COMObject, NetOffice.VisioApi.LPVISIOWINDOWS
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOWINDOWS);
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
                    _type = typeof(LPVISIOWINDOWS);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOWINDOWS() : base()
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
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.VisioApi.IVWindow this[Int16 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Count");
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
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVWindow get_ItemFromID(Int32 nID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ItemFromID", typeof(NetOffice.VisioApi.IVWindow), nID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemFromID")]
		public virtual NetOffice.VisioApi.IVWindow ItemFromID(Int32 nID)
		{
			return get_ItemFromID(nID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="captionOrIndex">object captionOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVWindow get_ItemEx(object captionOrIndex)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ItemEx", typeof(NetOffice.VisioApi.IVWindow), captionOrIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemEx
		/// </summary>
		/// <param name="captionOrIndex">object captionOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemEx")]
		public virtual NetOffice.VisioApi.IVWindow ItemEx(object captionOrIndex)
		{
			return get_ItemEx(captionOrIndex);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void VoidArrange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "VoidArrange");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption, nFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption, nFlags, nType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption, nFlags, nType, nLeft);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nArrangeFlags">optional object nArrangeFlags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Arrange(object nArrangeFlags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Arrange", nArrangeFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Arrange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Arrange");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		/// <param name="bstrMergeClass">optional object bstrMergeClass</param>
		/// <param name="nMergePosition">optional object nMergePosition</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass, object nMergePosition)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID, bstrMergeClass, nMergePosition });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption, nFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption, nFlags, nType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption, nFlags, nType, nLeft);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		/// <param name="bstrMergeClass">optional object bstrMergeClass</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID, bstrMergeClass });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVWindow>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVWindow>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVWindow>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVWindow>

        /// <summary>
        /// SupportByVersion Visio, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.VisioApi.IVWindow> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVWindow item in innerEnumerator)
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

