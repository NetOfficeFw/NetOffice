using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisUICtrlTypes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeBUTTON = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeBUTTON_OWNERDRAW = 33,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeOWNERDRAW_BUTTON = 33,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSPLITBUTTON = 17,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypePALETTEBUTTONNOMRU = 17,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSPLITBUTTON_MRU_COLOR = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypePALETTEBUTTON = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSPINBUTTON = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSPLITBUTTON_MRU_COMMAND = 18,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypePALETTEBUTTONICON = 18,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeEDITBOX = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeCOMBOBOX = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeCOMBOBOX_SORTED = 129,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>272</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeDROPDOWN = 272,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>273</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeDROPDOWN_SORTED = 273,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeDROPDOWN_OWNERDRAW = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeCOMBODRAW = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>257</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeDROPDOWN_SORTED_OWNERDRAW = 257,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeLABEL = 2048,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSWATCH = 32768,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32769</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSWATCH_COLORS = 32769,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeEND = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSTATE = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSTATE_BUTTON = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeHIERBUTTON = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSTATE_HIERBUTTON = 5,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeDROPBUTTON = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSTATE_DROPBUTTON = 9,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypePUSHBUTTON = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeLISTBOX = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>513</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeLISTBOXDRAW = 513,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeCOLORBOX = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeMESSAGE = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visCtrlTypeSPACER = 16384
	}
}