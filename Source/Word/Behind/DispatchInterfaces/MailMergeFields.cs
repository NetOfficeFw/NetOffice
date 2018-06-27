using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface MailMergeFields 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835151.aspx </remarks>
	public class MailMergeFields : COMObject, NetOffice.WordApi.MailMergeFields
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
                    _contractType = typeof(NetOffice.WordApi.MailMergeFields);
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
                    _type = typeof(MailMergeFields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public MailMergeFields() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195330.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834563.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840148.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837178.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.WordApi.MailMergeField this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "Item", typeof(NetOffice.WordApi.MailMergeField), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836028.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField Add(NetOffice.WordApi.Range range, string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "Add", typeof(NetOffice.WordApi.MailMergeField), range, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultAskText">optional object defaultAskText</param>
		/// <param name="askOnce">optional object askOnce</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText, object askOnce)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", typeof(NetOffice.WordApi.MailMergeField), new object[]{ range, name, prompt, defaultAskText, askOnce });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", typeof(NetOffice.WordApi.MailMergeField), range, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", typeof(NetOffice.WordApi.MailMergeField), range, name, prompt);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839897.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultAskText">optional object defaultAskText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddAsk(NetOffice.WordApi.Range range, string name, object prompt, object defaultAskText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddAsk", typeof(NetOffice.WordApi.MailMergeField), range, name, prompt, defaultAskText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultFillInText">optional object defaultFillInText</param>
		/// <param name="askOnce">optional object askOnce</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText, object askOnce)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", typeof(NetOffice.WordApi.MailMergeField), range, prompt, defaultFillInText, askOnce);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", typeof(NetOffice.WordApi.MailMergeField), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", typeof(NetOffice.WordApi.MailMergeField), range, prompt);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836431.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="prompt">optional object prompt</param>
		/// <param name="defaultFillInText">optional object defaultFillInText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddFillIn(NetOffice.WordApi.Range range, object prompt, object defaultFillInText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddFillIn", typeof(NetOffice.WordApi.MailMergeField), range, prompt, defaultFillInText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		/// <param name="falseAutoText">optional object falseAutoText</param>
		/// <param name="falseText">optional object falseText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText, object falseText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", typeof(NetOffice.WordApi.MailMergeField), new object[]{ range, mergeField, comparison, compareTo, trueAutoText, trueText, falseAutoText, falseText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", typeof(NetOffice.WordApi.MailMergeField), range, mergeField, comparison);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", typeof(NetOffice.WordApi.MailMergeField), range, mergeField, comparison, compareTo);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", typeof(NetOffice.WordApi.MailMergeField), new object[]{ range, mergeField, comparison, compareTo, trueAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", typeof(NetOffice.WordApi.MailMergeField), new object[]{ range, mergeField, comparison, compareTo, trueAutoText, trueText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845762.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		/// <param name="trueAutoText">optional object trueAutoText</param>
		/// <param name="trueText">optional object trueText</param>
		/// <param name="falseAutoText">optional object falseAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo, object trueAutoText, object trueText, object falseAutoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddIf", typeof(NetOffice.WordApi.MailMergeField), new object[]{ range, mergeField, comparison, compareTo, trueAutoText, trueText, falseAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195621.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddMergeRec(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddMergeRec", typeof(NetOffice.WordApi.MailMergeField), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839562.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddMergeSeq(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddMergeSeq", typeof(NetOffice.WordApi.MailMergeField), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837747.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddNext(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddNext", typeof(NetOffice.WordApi.MailMergeField), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836276.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddNextIf", typeof(NetOffice.WordApi.MailMergeField), range, mergeField, comparison, compareTo);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836276.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddNextIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddNextIf", typeof(NetOffice.WordApi.MailMergeField), range, mergeField, comparison);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="valueText">optional object valueText</param>
		/// <param name="valueAutoText">optional object valueAutoText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText, object valueAutoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSet", typeof(NetOffice.WordApi.MailMergeField), range, name, valueText, valueAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSet", typeof(NetOffice.WordApi.MailMergeField), range, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197837.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="name">string name</param>
		/// <param name="valueText">optional object valueText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddSet(NetOffice.WordApi.Range range, string name, object valueText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSet", typeof(NetOffice.WordApi.MailMergeField), range, name, valueText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841008.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		/// <param name="compareTo">optional object compareTo</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison, object compareTo)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSkipIf", typeof(NetOffice.WordApi.MailMergeField), range, mergeField, comparison, compareTo);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841008.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="mergeField">string mergeField</param>
		/// <param name="comparison">NetOffice.WordApi.Enums.WdMailMergeComparison comparison</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.MailMergeField AddSkipIf(NetOffice.WordApi.Range range, string mergeField, NetOffice.WordApi.Enums.WdMailMergeComparison comparison)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.MailMergeField>(this, "AddSkipIf", typeof(NetOffice.WordApi.MailMergeField), range, mergeField, comparison);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.MailMergeField>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.MailMergeField>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
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
        public virtual IEnumerator<NetOffice.WordApi.MailMergeField> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

