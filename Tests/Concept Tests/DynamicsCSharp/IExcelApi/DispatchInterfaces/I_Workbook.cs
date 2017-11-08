using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Duck;
using NetOffice.Attributes;

namespace NetOffice.IExcelApi
{
    [SyntaxBypass]
    public interface I_Workbook_ : ICOMObject
    {
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        object get_Colors(object index);

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        void set_Colors(object index, object value);

        [Redirect("get_Colors")]
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        object Colors(object index);
    }

    [EntityType(EntityType.IsDispatchInterface)]
    public interface I_Workbook : I_Workbook_
    {
        new object Colors { get; set; }

        string Name { get; }

        ISheets Sheets { get; }
        
        void SaveAs(object filename);
    }
}
