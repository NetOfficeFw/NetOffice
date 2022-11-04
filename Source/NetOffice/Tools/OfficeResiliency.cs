using System;
using System.IO;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Helper class to read data from the Microsoft Office resiliency registry values.
    /// Converts binary data to the <see cref="DisabledItemType"/> values.
    /// </summary>
    public class OfficeResiliency
    {
        private static readonly Encoding UTF16LE = Encoding.Unicode;

        /// <summary>
        /// Method for converting the binary data from the Resiliency\DisabledItems
        /// registry keys into the <see cref="DisabledItem"/> object.
        /// </summary>
        /// <param name="rawData">Binary data from the DisabledItems registry key.</param>
        /// <returns></returns>
        /// <remarks>
        /// DisabledItems registry keys are stored at <c>HKEY_CURRENT_USER\Software\Microsoft\Office\[OfficeAppVersion]\[OfficeAppName]\Resiliency\DisabledItems</c>
        /// as subkeys.
        /// The [OfficeAppVersion] is MS Office release version like <c>15.0</c> or <c>16.0</c>.
        /// The [OfficeAppName] is the name of a MS Office application like <c>Word</c>, <c>Excel</c>, <c>PowerPoint</c>, <c>Outlook</c> and others.
        /// </remarks>
        public static DisabledItem Parse(byte[] rawData)
        {
            using (var stream = new MemoryStream(rawData))
            using (var reader = new BinaryReader(stream))
            {
                var disabledItemTypeValue = reader.ReadInt32();
                var countData = reader.ReadInt32();
                var countExtraData = reader.ReadInt32();
                var offset = reader.BaseStream.Position;

                var disabledItemType = (DisabledItemType)disabledItemTypeValue;

                var item = new DisabledItem()
                {
                    DisabledItemType = disabledItemType
                };

                switch (disabledItemType)
                {
                    case DisabledItemType.AddInByFilename:
                    case DisabledItemType.AddInByDEPFilename:
                        
                        if (countData > 2)
                        {
                            var moduleBytes = reader.ReadBytes(countData - 2);
                            var module = UTF16LE.GetString(moduleBytes);
                            item.Module = module;
                        }

                        if (countExtraData > 2)
                        {
                            reader.BaseStream.Position = offset + countData;
                            var friendlyNameBytes = reader.ReadBytes(countExtraData - 2);
                            var friendlyName = UTF16LE.GetString(friendlyNameBytes);
                            item.FriendlyName = friendlyName;
                        }
                        break;
                }

                return item;
            }
        }
    }
}
