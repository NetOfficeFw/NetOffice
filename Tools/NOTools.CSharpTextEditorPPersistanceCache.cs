using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    public class PersistanceCache : IEnumerable<string>
    {
        internal PersistanceCache(CodeEditorControl parent)
        {
            Parent = parent;
        }

        private CodeEditorControl Parent { get; set; }

        public IEnumerator<string> GetEnumerator()
        {
            List<string> list = new List<string>();

            Dictionary<string, string> index = Parent.wpfControl1.CurrentFile.Persistence.LoadCacheIndex();
            foreach (var item in index)
            {
                int pos = item.Value.LastIndexOf(".");
                if (pos > -1)
                {
                    pos = item.Value.LastIndexOf(".", pos-1);
                    if (pos > -1)
                    {
                        string name = item.Value.Substring(0, pos);
                        if (!list.Contains(name))
                            list.Add(name);
                    }
                }
                else
                {
                    if (!list.Contains(item.Value))
                        list.Add(item.Value);
                }
            }
            foreach (var item in list)
            {
                yield return item;
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
