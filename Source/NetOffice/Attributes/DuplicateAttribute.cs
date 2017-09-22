using System;

namespace NetOffice.Attributes
{
    /*
        Duplicate Attribute means an equal type id in another type library.
        There are also duplicates in the same type library but we dont care.

        (See list of known duplicates below.)
    */

    /// <summary>
    /// Known duplicate in MS-Office Automation Model
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class DuplicateAttribute : System.Attribute
    {
        /// <summary>
        /// Duplicate Type Name (comma separated, if more than 1 duplicate)
        /// </summary>
        public readonly string To;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="to">duplicate type name</param>
        public DuplicateAttribute(string to)
        {
            To = to;
        }
    }
}

/*
* 
*  Known duplicate types by type-id in another type library
*  ----------------------------------
* 
Office.Adjustments			Excel.Adjustments		    000C0310-0000-0000-C000-000000000046
Office.CalloutFormat		Excel.CalloutFormat		    000C0311-0000-0000-C000-000000000046
Office.ColorFormat			Excel.ColorFormat		    000C0312-0000-0000-C000-000000000046
Office.FillFormat			Excel.FillFormat		    000C0314-0000-0000-C000-000000000046
Office.LineFormat			Excel.LineFormat		    000C0317-0000-0000-C000-000000000046
Office.ShapeNode			Excel.ShapeNode			    000C0318-0000-0000-C000-000000000046
Office.ShapeNodes			Excel.ShapeNodes		    000C0319-0000-0000-C000-000000000046
Office.PictureFormat		Excel.PictureFormat		    000C031A-0000-0000-C000-000000000046
Office.ShadowFormat			Excel.ShadowFormat		    000C031B-0000-0000-C000-000000000046
Office.TextEffectFormat		Excel.TextEffectFormat		000C031F-0000-0000-C000-000000000046
Office.ThreeDFormat			Excel.ThreeDFormat		    000C0321-0000-0000-C000-000000000046
Office.DiagramNodes			Excel.DiagramNodes		    000C036E-0000-0000-C000-000000000046
Office.DiagramNodeChildren	Excel.DiagramNodeChildren	000C036F-0000-0000-C000-000000000046
Office.DiagramNode			Excel.DiagramNode		    000C0370-0000-0000-C000-000000000046
Office.TextFrame2			Excel.TextFrame2		    000C0398-0000-0000-C000-000000000046

* Known duplicate types by type-id in same type library
*  ----------------------------------
* 
Access.Form	            Access.FormOld	                483615A0-74BE-101B-AF4E-00AA003F0F07
Access.Form	            Access.FormOldV10	            483615A0-74BE-101B-AF4E-00AA003F0F07
Access.Report	        Access.ReportOld	            27CE30A0-91FF-101B-AF4E-00AA003F0F07
Access.Report	        Access.ReportOldV10	            27CE30A0-91FF-101B-AF4E-00AA003F0F07
Access.FormOld	        Access.Form	                    483615A0-74BE-101B-AF4E-00AA003F0F07
Access.ReportOld	    Access.Report	                27CE30A0-91FF-101B-AF4E-00AA003F0F07
Access.FormOldV10	    Access.Form	                    483615A0-74BE-101B-AF4E-00AA003F0F08
Access.ReportOldV10	    Access.Report	                ECD1EADA-D373-11D3-8D21-0050048383FB
ADODB.Command15	        ADODB.Command15_Deprecated	    00000508-0000-0010-8000-00AA006D2EA4
ADODB._Connection	    ADODB._Connection_Deprecated	00000550-0000-0010-8000-00AA006D2EA4
ADODB.Connection15	    ADODB.Connection15_Deprecated	00000515-0000-0010-8000-00AA006D2EA4
ADODB._Recordset	    ADODB.Recordset21	            00000555-0000-0010-8000-00AA006D2EA4
ADODB._Recordset	    ADODB._Recordset_Deprecated	    00000555-0000-0010-8000-00AA006D2EA4
ADODB._Recordset	    ADODB.Recordset21_Deprecated	00000555-0000-0010-8000-00AA006D2EA4
ADODB.Recordset20	    ADODB.Recordset20_Deprecated	0000054F-0000-0010-8000-00AA006D2EA4
ADODB.Recordset15	    ADODB.Recordset15_Deprecated	0000050E-0000-0010-8000-00AA006D2EA4
ADODB.Fields	        ADODB.Fields20	                0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Fields	        ADODB.Fields_Deprecated	        0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Fields	        ADODB.Fields20_Deprecated	    0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Fields15	        ADODB.Fields15_Deprecated	    00000506-0000-0010-8000-00AA006D2EA4
ADODB.Field	            ADODB.Field20	                0000054C-0000-0010-8000-00AA006D2EA4
ADODB.Field	            ADODB.Field_Deprecated	        0000054C-0000-0010-8000-00AA006D2EA4
ADODB.Field	A           DODB.Field20_Deprecated	        0000054C-0000-0010-8000-00AA006D2EA4
ADODB._Parameter	    ADODB._Parameter_Deprecated	    0000050C-0000-0010-8000-00AA006D2EA4
ADODB.Parameters	    ADODB.Parameters_Deprecated	    0000050D-0000-0010-8000-00AA006D2EA4
ADODB._Command	        ADODB._Command_Deprecated	        0000054E-0000-0010-8000-00AA006D2EA4
ADODB.ConnectionEvents	ADODB.ConnectionEvents_Deprecated	00000400-0000-0010-8000-00AA006D2EA4
ADODB.RecordsetEvents	ADODB.RecordsetEvents_Deprecated	00000266-0000-0010-8000-00AA006D2EA4
ADODB.Field15	        ADODB.Field15_Deprecated	        00000505-0000-0010-8000-00AA006D2EA4
ADODB.Recordset21	    ADODB._Recordset	            00000555-0000-0010-8000-00AA006D2EA4
ADODB.Recordset21	    ADODB.Recordset21_Deprecated	00000555-0000-0010-8000-00AA006D2EA4
ADODB.Fields20	        ADODB.Fields	                    0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Fields20	        ADODB.Fields20_Deprecated	        0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Field20	        ADODB.Field	                        0000054C-0000-0010-8000-00AA006D2EA4
ADODB.Field20	        ADODB.Field20_Deprecated	        0000054C-0000-0010-8000-00AA006D2EA4
ADODB._Record	        ADODB._Record_Deprecated	        00000562-0000-0010-8000-00AA006D2EA4
ADODB._Stream	        ADODB._Stream_Deprecated	        00000565-0000-0010-8000-00AA006D2EA4
ADODB.Command15_Deprecated	    ADODB.Command15	        00000508-0000-0010-8000-00AA006D2EA4
ADODB._Connection_Deprecated	ADODB._Connection	    00000550-0000-0010-8000-00AA006D2EA4
ADODB.Connection15_Deprecated	ADODB.Connection15	    00000515-0000-0010-8000-00AA006D2EA4
ADODB._Recordset_Deprecated	    ADODB._Recordset	    00000556-0000-0010-8000-00AA006D2EA4
ADODB.Recordset21_Deprecated	ADODB._Recordset	    00000555-0000-0010-8000-00AA006D2EA4
ADODB.Recordset21_Deprecated	ADODB.Recordset21	    00000555-0000-0010-8000-00AA006D2EA4
ADODB.Recordset20_Deprecated	ADODB.Recordset20	    0000054F-0000-0010-8000-00AA006D2EA4
ADODB.Recordset15_Deprecated	ADODB.Recordset15	    0000050E-0000-0010-8000-00AA006D2EA4
ADODB.Fields_Deprecated	        ADODB.Fields	        00000564-0000-0010-8000-00AA006D2EA4
ADODB.Fields20_Deprecated	    ADODB.Fields	        0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Fields20_Deprecated	    ADODB.Fields20	        0000054D-0000-0010-8000-00AA006D2EA4
ADODB.Fields15_Deprecated	    ADODB.Fields15	        00000506-0000-0010-8000-00AA006D2EA4
ADODB.Field_Deprecated	        ADODB.Field	            00000569-0000-0010-8000-00AA006D2EA4
ADODB.Field20_Deprecated	    ADODB.Field	            0000054C-0000-0010-8000-00AA006D2EA4
ADODB.Field20_Deprecated	    ADODB.Field20	        0000054C-0000-0010-8000-00AA006D2EA4
ADODB._Parameter_Deprecated	    ADODB._Parameter	    0000050C-0000-0010-8000-00AA006D2EA4
ADODB.Parameters_Deprecated	    ADODB.Parameters	    0000050D-0000-0010-8000-00AA006D2EA4
ADODB._Command_Deprecated	    ADODB._Command	        0000054E-0000-0010-8000-00AA006D2EA4
ADODB.ConnectionEvents_Deprecated	ADODB.ConnectionEvents	00000400-0000-0010-8000-00AA006D2EA4
ADODB.RecordsetEvents_Deprecated	ADODB.RecordsetEvents	00000266-0000-0010-8000-00AA006D2EA4
ADODB._Record_Deprecated	    ADODB._Record	        00000562-0000-0010-8000-00AA006D2EA4
ADODB._Stream_Deprecated	    ADODB._Stream	        00000565-0000-0010-8000-00AA006D2EA4
ADODB.Field15_Deprecated	    ADODB.Field15	        00000505-0000-0010-8000-00AA006D2EA4
ADODB.ConnectionEventsVt	    ADODB.ConnectionEventsVt_Deprecated	        00000402-0000-0010-8000-00AA006D2EA4
ADODB.RecordsetEventsVt	        ADODB.RecordsetEventsVt_Deprecated	        00000403-0000-0010-8000-00AA006D2EA4
ADODB.ADORecordsetConstruction	ADODB.ADORecordsetConstruction_Deprecated	00000283-0000-0010-8000-00AA006D2EA4
ADODB.ConnectionEventsVt_Deprecated	        ADODB.ConnectionEventsVt	    00000402-0000-0010-8000-00AA006D2EA4
ADODB.RecordsetEventsVt_Deprecated	        ADODB.RecordsetEventsVt	        00000403-0000-0010-8000-00AA006D2EA4
ADODB.ADORecordsetConstruction_Deprecated	ADODB.ADORecordsetConstruction	00000283-0000-0010-8000-00AA006D2EA4

*  Known duplicate types by type-id in a another type library we dont spend a wrapper for
*  ----------------------------------
MSForms.Font	stdole.Font	    BEF6E003-A874-101A-8BBA-00AA00300CAB
MSForms.IFont	stdole.IFont	BEF6E002-A874-101A-8BBA-00AA00300CAB
* 
*/