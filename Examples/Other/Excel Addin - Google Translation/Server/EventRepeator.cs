using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.Server
{
    public class DataEventRepeator : MarshalByRefObject
    {
        public event TranslationEventHandler Translation;

        public void OnTranslation(TranslateOperationResult result)
        {
            if (Translation != null)
                Translation(result);
        }
    }

    public class DataEventRepeators : List<DataEventRepeator>
    {
    }
}
