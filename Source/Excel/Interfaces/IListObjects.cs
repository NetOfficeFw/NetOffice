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
	/// Interface IListObjects 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class IListObjects : COMObject, IEnumerableProvider<NetOffice.ExcelApi.ListObject>
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
                    _type = typeof(IListObjects);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IListObjects(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IListObjects(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListObjects(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListObjects(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListObjects(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListObjects(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListObjects() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListObjects(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.ListObject this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObject>(this, "_Default", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		/// <param name="xlListObjectHasHeaders">optional NetOffice.ExcelApi.Enums.XlYesNoGuess XlListObjectHasHeaders = 0</param>
		/// <param name="destination">optional object destination</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add(object sourceType, object source, object linkSource, object xlListObjectHasHeaders, object destination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, new object[]{ sourceType, source, linkSource, xlListObjectHasHeaders, destination });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		/// <param name="xlListObjectHasHeaders">optional NetOffice.ExcelApi.Enums.XlYesNoGuess XlListObjectHasHeaders = 0</param>
		/// <param name="destination">optional object destination</param>
		/// <param name="tableStyleName">optional object tableStyleName</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add(object sourceType, object source, object linkSource, object xlListObjectHasHeaders, object destination, object tableStyleName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, new object[]{ sourceType, source, linkSource, xlListObjectHasHeaders, destination, tableStyleName });
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add(object sourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add(object sourceType, object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType, source);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add(object sourceType, object source, object linkSource)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType, source, linkSource);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		/// <param name="xlListObjectHasHeaders">optional NetOffice.ExcelApi.Enums.XlYesNoGuess XlListObjectHasHeaders = 0</param>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject Add(object sourceType, object source, object linkSource, object xlListObjectHasHeaders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType, source, linkSource, xlListObjectHasHeaders);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		/// <param name="xlListObjectHasHeaders">optional NetOffice.ExcelApi.Enums.XlYesNoGuess XlListObjectHasHeaders = 0</param>
		/// <param name="destination">optional object destination</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject _Add(object sourceType, object source, object linkSource, object xlListObjectHasHeaders, object destination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "_Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, new object[]{ sourceType, source, linkSource, xlListObjectHasHeaders, destination });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject _Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "_Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject _Add(object sourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "_Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject _Add(object sourceType, object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "_Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType, source);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject _Add(object sourceType, object source, object linkSource)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "_Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType, source, linkSource);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">optional NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType = 1</param>
		/// <param name="source">optional object source</param>
		/// <param name="linkSource">optional object linkSource</param>
		/// <param name="xlListObjectHasHeaders">optional NetOffice.ExcelApi.Enums.XlYesNoGuess XlListObjectHasHeaders = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ListObject _Add(object sourceType, object source, object linkSource, object xlListObjectHasHeaders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ListObject>(this, "_Add", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType, sourceType, source, linkSource, xlListObjectHasHeaders);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.ListObject>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.ListObject>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.ListObject>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.ListObject>

        /// <summary>
        /// SupportByVersion Excel, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.ListObject> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.ListObject item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}