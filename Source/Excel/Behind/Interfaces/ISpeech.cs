using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// Interface ISpeech 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class ISpeech : COMObject, NetOffice.ExcelApi.ISpeech
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
                    _contractType = typeof(NetOffice.ExcelApi.ISpeech);
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
                    _type = typeof(ISpeech);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISpeech() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlSpeakDirection Direction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSpeakDirection>(this, "Direction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool SpeakCellOnEnter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpeakCellOnEnter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpeakCellOnEnter", value);
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
		public virtual Int32 Speak(string text, object speakAsync, object speakXML, object purge)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak", text, speakAsync, speakXML, purge);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 Speak(string text)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak", text);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 Speak(string text, object speakAsync)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak", text, speakAsync);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="speakAsync">optional object speakAsync</param>
		/// <param name="speakXML">optional object speakXML</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 Speak(string text, object speakAsync, object speakXML)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Speak", text, speakAsync, speakXML);
		}

		#endregion

		#pragma warning restore
	}
}

