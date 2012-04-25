using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.AddinGuard
{
    public enum NotifyKind
    {
        Nothing = 0,

        AddinValuesIncrement = 8,
        AddinValuesDecrement = 9,

        AddinLoadBehaviorRestored = 1,
        AddinValueNameIsChanged = 2,
        AddinValueKindIsChanged = 3,
        AddinValueIsChanged = 4,

        DisabledItemNew = 10,
        DisabledItemDelete = 11,

        AddinSubKeysIncrement = 5,
        AddinSubKeysDecrement = 6,
        AddinSubKeyNameChanged = 7
    }
}
