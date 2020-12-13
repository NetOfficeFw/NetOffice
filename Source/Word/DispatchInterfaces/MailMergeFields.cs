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
	/// DispatchInterface MailMergeFields 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.mailmergefields"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class MailMergeFields : COMObject, IEnumerableProvider<NetOffice.WordApi.MailMergeField>
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
                    _type = typeof(MailMergeFields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MailMergeFields(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MailMergeFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MailMergeFields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.MailMergeField this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "Item", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField Add(NetOffice.WordApi.Range range, string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "Add", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddAsk"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultAskText">optional object defaultAskText</param>
		/// <param name="askOnce">optional object askOnce</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText, object askOnce)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, new object[]{ range, name, prompt, defaultAskText, askOnce });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddAsk"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddAsk"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name, prompt);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddAsk"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultAskText">optional object defaultAskText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name, prompt, defaultAskText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddFillIn"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultFillInText">optional object defaultFillInText</param>
		/// <param name="askOnce">optional object askOnce</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText, object askOnce)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, prompt, defaultFillInText, askOnce);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddFillIn"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddFillIn"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, prompt);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddFillIn"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultFillInText">optional object defaultFillInText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, prompt, defaultFillInText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		/// <param name="falseAutoText">optional object falseAutoText</param>
		/// <param name="falseText">optional object falseText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText, object falseText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, new object[]{ range, mergeField, comparison, compareTo, trueAutoText, trueText, falseAutoText, falseText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, mergeField, comparison);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, mergeField, comparison, compareTo);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, new object[]{ range, mergeField, comparison, compareTo, trueAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, new object[]{ range, mergeField, comparison, compareTo, trueAutoText, trueText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		/// <param name="falseAutoText">optional object falseAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, new object[]{ range, mergeField, comparison, compareTo, trueAutoText, trueText, falseAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddMergeRec"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddMergeRec(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddMergeRec", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddMergeSeq"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddMergeSeq(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddMergeSeq", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddNext"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddNext(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddNext", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddNextIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddNextIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, mergeField, comparison, compareTo);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddNextIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddNextIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, mergeField, comparison);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddSet"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="valueText">optional object valueText</param>
		/// <param name="valueAutoText">optional object valueAutoText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText, object valueAutoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSet", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name, valueText, valueAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddSet"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSet", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddSet"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="valueText">optional object valueText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSet", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, name, valueText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddSkipIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSkipIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, mergeField, comparison, compareTo);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.MailMergeFields.AddSkipIf"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSkipIf", NetOffice.WordApi.MailMergeField.LateBindingApiWrapperType, range, mergeField, comparison);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.MailMergeField>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.MailMergeField>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.MailMergeField>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.MailMergeField>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.MailMergeField> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.MailMergeField item in innerEnumerator)
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