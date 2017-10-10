using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVWindows 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class IVWindows : COMObject, IEnumerableProvider<NetOffice.VisioApi.IVWindow>
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
                    _type = typeof(IVWindows);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVWindows(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVWindows(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindows(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindows(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindows(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindows(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindows() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVWindows(string progId) : base(progId)
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
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
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
		public NetOffice.VisioApi.IVWindow this[Int16 index]
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVWindow get_ItemFromID(Int32 nID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ItemFromID", NetOffice.VisioApi.IVWindow.LateBindingApiWrapperType, nID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemFromID")]
		public NetOffice.VisioApi.IVWindow ItemFromID(Int32 nID)
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
		public NetOffice.VisioApi.IVWindow get_ItemEx(object captionOrIndex)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "ItemEx", NetOffice.VisioApi.IVWindow.LateBindingApiWrapperType, captionOrIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemEx
		/// </summary>
		/// <param name="captionOrIndex">object captionOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemEx")]
		public NetOffice.VisioApi.IVWindow ItemEx(object captionOrIndex)
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
		public void VoidArrange()
		{
			 Factory.ExecuteMethod(this, "VoidArrange");
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
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption);
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
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption, nFlags);
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
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption, nFlags, nType);
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
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", bstrCaption, nFlags, nType, nLeft);
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
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop });
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
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add_WithoutMergeArgs", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nArrangeFlags">optional object nArrangeFlags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Arrange(object nArrangeFlags)
		{
			 Factory.ExecuteMethod(this, "Arrange", nArrangeFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Arrange()
		{
			 Factory.ExecuteMethod(this, "Arrange");
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass, object nMergePosition)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID, bstrMergeClass, nMergePosition });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption, nFlags);
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption, nFlags, nType);
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", bstrCaption, nFlags, nType, nLeft);
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop });
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth });
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight });
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID });
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
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "Add", new object[]{ bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID, bstrMergeClass });
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
        public IEnumerator<NetOffice.VisioApi.IVWindow> GetEnumerator()
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