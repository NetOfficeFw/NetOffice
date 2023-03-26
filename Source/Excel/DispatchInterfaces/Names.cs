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
	/// DispatchInterface Names 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "_Default")]
	public class Names : COMObject, IEnumerableProvider<NetOffice.ExcelApi.Name>
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
                    _type = typeof(Names);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Names(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Names(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Names(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Application"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Creator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Count"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		/// <param name="refersToR1C1">optional object refersToR1C1</param>
		/// <param name="refersToR1C1Local">optional object refersToR1C1Local</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1, refersToR1C1Local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, name, refersTo);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, name, refersTo, visible);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, name, refersTo, visible, macroType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey, category });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Names.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		/// <param name="refersToR1C1">optional object refersToR1C1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, new object[]{ name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1 });
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Custom Indexer
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public NetOffice.ExcelApi.Name this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "_Default", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, index);
			}
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Custom Indexer
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="indexLocal">optional object indexLocal</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public NetOffice.ExcelApi.Name this[object index, object indexLocal]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "_Default", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, index, indexLocal);
			}
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="indexLocal">optional object indexLocal</param>
        /// <param name="refersTo">optional object refersTo</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.Name this[object index, object indexLocal, object refersTo]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "_Default", NetOffice.ExcelApi.Name.LateBindingApiWrapperType, index, indexLocal, refersTo);
			}
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Name>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Name>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Name>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Name>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.ExcelApi.Name> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.ExcelApi.Name item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable
       
		/// <summary>
		/// SupportByVersion Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}