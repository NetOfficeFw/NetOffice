using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ISlicers 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class ISlicers : COMObject, IEnumerableProvider<NetOffice.ExcelApi.Slicer>
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
                    _type = typeof(ISlicers);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISlicers(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISlicers(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.Slicer this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicer>(this, "_Default", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, new object[]{ slicerDestination, level, name, caption, top, left, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, slicerDestination);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, slicerDestination, level);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, slicerDestination, level, name);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, slicerDestination, level, name, caption);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, new object[]{ slicerDestination, level, name, caption, top });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, new object[]{ slicerDestination, level, name, caption, top, left });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="slicerDestination">object slicerDestination</param>
		/// <param name="level">optional object level</param>
		/// <param name="name">optional object name</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="top">optional object top</param>
		/// <param name="left">optional object left</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Slicer>(this, "Add", NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType, new object[]{ slicerDestination, level, name, caption, top, left, width });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Slicer>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Slicer>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Slicer>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Slicer>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.Slicer> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Slicer item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}