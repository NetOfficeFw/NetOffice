using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface AutoCorrect 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839324.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AutoCorrect : COMObject
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
                    _type = typeof(AutoCorrect);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AutoCorrect(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AutoCorrect(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AutoCorrect(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196916.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196945.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822117.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836049.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectDays
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectDays");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectDays", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196887.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectInitialCaps
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectInitialCaps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectInitialCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194989.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectSentenceCaps
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectSentenceCaps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectSentenceCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194587.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ReplaceText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReplaceText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReplaceText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839917.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoCorrectEntries Entries
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrectEntries>(this, "Entries", NetOffice.WordApi.AutoCorrectEntries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193733.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FirstLetterExceptions FirstLetterExceptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FirstLetterExceptions>(this, "FirstLetterExceptions", NetOffice.WordApi.FirstLetterExceptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191750.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool FirstLetterAutoAdd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FirstLetterAutoAdd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstLetterAutoAdd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837875.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TwoInitialCapsExceptions TwoInitialCapsExceptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TwoInitialCapsExceptions>(this, "TwoInitialCapsExceptions", NetOffice.WordApi.TwoInitialCapsExceptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834297.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool TwoInitialCapsAutoAdd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TwoInitialCapsAutoAdd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TwoInitialCapsAutoAdd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192775.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectCapsLock
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectCapsLock");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectCapsLock", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837157.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectHangulAndAlphabet
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectHangulAndAlphabet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectHangulAndAlphabet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836549.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.HangulAndAlphabetExceptions HangulAndAlphabetExceptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HangulAndAlphabetExceptions>(this, "HangulAndAlphabetExceptions", NetOffice.WordApi.HangulAndAlphabetExceptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839148.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool HangulAndAlphabetAutoAdd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HangulAndAlphabetAutoAdd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HangulAndAlphabetAutoAdd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822942.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ReplaceTextFromSpellingChecker
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReplaceTextFromSpellingChecker");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReplaceTextFromSpellingChecker", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836260.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool OtherCorrectionsAutoAdd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OtherCorrectionsAutoAdd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherCorrectionsAutoAdd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197197.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.OtherCorrectionsExceptions OtherCorrectionsExceptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OtherCorrectionsExceptions>(this, "OtherCorrectionsExceptions", NetOffice.WordApi.OtherCorrectionsExceptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192767.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectKeyboardSetting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectKeyboardSetting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectKeyboardSetting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822882.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool CorrectTableCells
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectTableCells");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectTableCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821332.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisplayAutoCorrectOptions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAutoCorrectOptions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAutoCorrectOptions", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
