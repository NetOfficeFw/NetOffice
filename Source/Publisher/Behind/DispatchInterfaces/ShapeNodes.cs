using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface ShapeNodes 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	public class ShapeNodes : COMObject, NetOffice.PublisherApi.ShapeNodes
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
                    _contractType = typeof(NetOffice.PublisherApi.ShapeNodes);
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
                    _type = typeof(ShapeNodes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ShapeNodes() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.PublisherApi.ShapeNode this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeNode>(this, "Item", typeof(NetOffice.PublisherApi.ShapeNode), index);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Delete(Int32 index)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", index);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		/// <param name="y2">optional object y2</param>
		/// <param name="x3">optional object x3</param>
		/// <param name="y3">optional object y3</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2, object y2, object x3, object y3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ index, segmentType, editingType, x1, y1, x2, y2, x3, y3 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ index, segmentType, editingType, x1, y1 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ index, segmentType, editingType, x1, y1, x2 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		/// <param name="y2">optional object y2</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2, object y2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ index, segmentType, editingType, x1, y1, x2, y2 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		/// <param name="x2">optional object x2</param>
		/// <param name="y2">optional object y2</param>
		/// <param name="x3">optional object x3</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Insert(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType, NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1, object x2, object y2, object x3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", new object[]{ index, segmentType, editingType, x1, y1, x2, y2, x3 });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetEditingType(Int32 index, NetOffice.OfficeApi.Enums.MsoEditingType editingType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetEditingType", index, editingType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetPosition(Int32 index, object x1, object y1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetPosition", index, x1, y1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="segmentType">NetOffice.OfficeApi.Enums.MsoSegmentType segmentType</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SetSegmentType(Int32 index, NetOffice.OfficeApi.Enums.MsoSegmentType segmentType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetSegmentType", index, segmentType);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.ShapeNode>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.ShapeNode>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.ShapeNode>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.ShapeNode>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public virtual IEnumerator<NetOffice.PublisherApi.ShapeNode> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.ShapeNode item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

