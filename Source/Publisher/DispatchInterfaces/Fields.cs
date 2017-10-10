using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Fields 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Fields : COMObject , NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Field>
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
                    _type = typeof(Fields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Fields(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Fields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Fields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PublisherApi.Field this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "Item", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Unlink()
		{
			 Factory.ExecuteMethod(this, "Unlink");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddHorizontalInVertical(NetOffice.PublisherApi.TextRange range, string text)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "AddHorizontalInVertical", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, range, text);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="fontSize">optional object FontSize = 10</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise, object fontName, object fontSize)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "AddPhoneticGuide", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, new object[]{ range, text, alignment, raise, fontName, fontSize });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "AddPhoneticGuide", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, range, text);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "AddPhoneticGuide", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, range, text, alignment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "AddPhoneticGuide", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, range, text, alignment, raise);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.PublisherApi.Enums.PbPhoneticGuideAlignmentType Alignment = 0</param>
		/// <param name="raise">optional object Raise = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Field AddPhoneticGuide(NetOffice.PublisherApi.TextRange range, string text, object alignment, object raise, object fontName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Field>(this, "AddPhoneticGuide", NetOffice.PublisherApi.Field.LateBindingApiWrapperType, new object[]{ range, text, alignment, raise, fontName });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.Field>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.Field>.GetComObjectEnumerator(ICOMObject parent)
        {
            return this;
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.Field>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Field item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Field> Member

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        public IEnumerator<NetOffice.PublisherApi.Field> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Field item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        [CustomEnumerator]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            int count = Count;
            object[] enumeratorObjects = new object[count];
            for (int i = 0; i < count; i++)
                enumeratorObjects[i] = this[i + 1];

            foreach (object item in enumeratorObjects)
                yield return item;
        }

        #endregion

        #pragma warning restore
    }
}