using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Subdocuments 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.subdocuments"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Subdocuments : COMObject, IEnumerableProvider<NetOffice.WordApi.Subdocument>
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
                    _type = typeof(Subdocuments);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Subdocuments(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Subdocuments(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Subdocuments(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Subdocuments(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Subdocuments(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Subdocuments(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Subdocuments() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Subdocuments(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.subdocuments.parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Expanded"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Expanded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Expanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Expanded", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.Subdocument this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "Item", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions, object readOnly, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, new object[]{ name, confirmConversions, readOnly, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, name, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, name, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions, object readOnly, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, name, confirmConversions, readOnly, passwordDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions, object readOnly, object passwordDocument, object passwordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, new object[]{ name, confirmConversions, readOnly, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions, object readOnly, object passwordDocument, object passwordTemplate, object revert)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, new object[]{ name, confirmConversions, readOnly, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromFile"/> </remarks>
		/// <param name="name">object name</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromFile(object name, object confirmConversions, object readOnly, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromFile", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, new object[]{ name, confirmConversions, readOnly, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.AddFromRange"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocument AddFromRange(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Subdocument>(this, "AddFromRange", NetOffice.WordApi.Subdocument.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Merge"/> </remarks>
		/// <param name="firstSubdocument">optional object firstSubdocument</param>
		/// <param name="lastSubdocument">optional object lastSubdocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Merge(object firstSubdocument, object lastSubdocument)
		{
			 Factory.ExecuteMethod(this, "Merge", firstSubdocument, lastSubdocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Merge"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Merge()
		{
			 Factory.ExecuteMethod(this, "Merge");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Merge"/> </remarks>
		/// <param name="firstSubdocument">optional object firstSubdocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Merge(object firstSubdocument)
		{
			 Factory.ExecuteMethod(this, "Merge", firstSubdocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Delete"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Subdocuments.Select"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Subdocument>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Subdocument>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Subdocument>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Subdocument>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Subdocument> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Subdocument item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}