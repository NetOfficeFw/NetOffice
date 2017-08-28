using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ISpeech 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class ISpeech : COMObject
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
                    _type = typeof(ISpeech);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISpeech(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISpeech(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpeech(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpeech(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpeech(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpeech(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpeech() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISpeech(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSpeakDirection Direction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSpeakDirection>(this, "Direction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool SpeakCellOnEnter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpeakCellOnEnter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpeakCellOnEnter", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		/// <param name="purge">optional object purge</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Speak(string text, object speakAsync, object speakXML, object purge)
		{
			return Factory.ExecuteInt32MethodGet(this, "Speak", text, speakAsync, speakXML, purge);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Speak(string text)
		{
			return Factory.ExecuteInt32MethodGet(this, "Speak", text);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Speak(string text, object speakAsync)
		{
			return Factory.ExecuteInt32MethodGet(this, "Speak", text, speakAsync);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Speak(string text, object speakAsync, object speakXML)
		{
			return Factory.ExecuteInt32MethodGet(this, "Speak", text, speakAsync, speakXML);
		}

		#endregion

		#pragma warning restore
	}
}
