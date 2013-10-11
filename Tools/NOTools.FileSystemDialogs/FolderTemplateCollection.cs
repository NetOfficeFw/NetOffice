using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class FolderTemplateCollection : List<FolderTemplate>
    {
    }
}
