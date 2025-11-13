# SensitivityLabel Implementation Plan - Issue #441

**Issue**: [#441 - SensitivityLabel missing](https://github.com/NetOfficeFw/NetOffice/issues/441)
**Pull Request**: [#445 - Implement SensitivityLabel object model](https://github.com/NetOfficeFw/NetOffice/pull/445)
**Milestone**: 1.9.8
**Status**: Implementation Complete - Awaiting Testing

## Overview

Implement the Microsoft Office `SensitivityLabel` API in NetOffice to support sensitivity labels for documents, workbooks, and presentations. This feature requires Office 2016+ and appropriate organizational licensing.

## API Documentation References

- **Office.SensitivityLabel**: https://learn.microsoft.com/en-us/office/vba/api/office.sensitivitylabel
- **Office.LabelInfo**: https://learn.microsoft.com/en-us/office/vba/api/office.labelinfo
- **Word.Document.SensitivityLabel**: https://learn.microsoft.com/en-us/office/vba/api/word.document.sensitivitylabel
- **Excel.Workbook.SensitivityLabel**: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.sensitivitylabel
- **PowerPoint.Presentation.SensitivityLabel**: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.sensitivitylabel

## Implementation Status

### ✅ Completed

#### 1. Core Office API Classes (PR #445)

**SensitivityLabel Class**
- **File**: `Source/Office/DispatchInterfaces/SensitivityLabel.cs`
- **Created**: 2025-11-05
- **Methods**:
  - `CreateLabelInfo()` - Creates a new LabelInfo object
  - `GetLabel()` - Gets current label information
  - `SetLabel(LabelInfo, object)` - Sets label information

**LabelInfo Class**
- **File**: `Source/Office/DispatchInterfaces/LabelInfo.cs`
- **Created**: 2025-11-05
- **Properties**:
  - `ActionId` (string) - GUID identifying the action
  - `AssignmentMethod` (MsoAssignmentMethod enum) - How label was assigned
  - `ContentBits` (int) - Content markings value
  - `IsEnabled` (bool) - Whether label is enabled
  - `Justification` (string) - Required when downgrading labels
  - `LabelId` (string) - GUID of sensitivity label
  - `LabelName` (string) - Display name of label
  - `SetDate` (DateTime) - Date when label was set
  - `SiteId` (string) - GUID of SharePoint site

**MsoAssignmentMethod Enum**
- **File**: `Source/Office/Enums/MsoAssignmentMethod.cs`
- **Created**: 2025-11-05
- **Values**:
  - `NOT_SET` (-1) - Assignment method not set
  - `STANDARD` (0) - Label applied by default
  - `PRIVILEGED` (1) - Label manually selected
  - `AUTO` (2) - Label applied automatically

#### 2. Document Object Properties (2025-11-13)

**Word.Document.SensitivityLabel**
- **File**: `Source/Word/DispatchInterfaces/_Document.cs:3523`
- **Type**: Read-only property
- **Returns**: `NetOffice.OfficeApi.SensitivityLabel`
- **Support**: Word 16+

**Excel.Workbook.SensitivityLabel**
- **File**: `Source/Excel/DispatchInterfaces/_Workbook.cs:2359`
- **Type**: Read-only property
- **Returns**: `NetOffice.OfficeApi.SensitivityLabel`
- **Support**: Excel 16+

**PowerPoint.Presentation.SensitivityLabel**
- **File**: `Source/PowerPoint/DispatchInterfaces/_Presentation.cs:1296`
- **Type**: Read-only property
- **Returns**: `NetOffice.OfficeApi.SensitivityLabel`
- **Support**: PowerPoint 16+

## Code Examples

### Example 1: Get Sensitivity Label (Word)

```csharp
using System;
using Word = NetOffice.WordApi;
using Office = NetOffice.OfficeApi;

// Open Word document
Word.Application wordApp = new Word.Application();
Word.Document doc = wordApp.Documents.Open(@"C:\path\to\document.docx");

// Get sensitivity label
Office.SensitivityLabel label = doc.SensitivityLabel;
if (label != null)
{
    Office.LabelInfo labelInfo = label.GetLabel();
    if (labelInfo != null)
    {
        Console.WriteLine($"Label Name: {labelInfo.LabelName}");
        Console.WriteLine($"Label ID: {labelInfo.LabelId}");
        Console.WriteLine($"Assignment Method: {labelInfo.AssignmentMethod}");
        Console.WriteLine($"Set Date: {labelInfo.SetDate}");
        Console.WriteLine($"Is Enabled: {labelInfo.IsEnabled}");
    }
}

// Cleanup
doc.Close();
wordApp.Quit();
wordApp.Dispose();
```

### Example 2: Set Sensitivity Label (Excel)

```csharp
using System;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;

// Open Excel workbook
Excel.Application excelApp = new Excel.Application();
Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\path\to\workbook.xlsx");

// Get sensitivity label
Office.SensitivityLabel label = workbook.SensitivityLabel;
if (label != null)
{
    // Create new label info
    Office.LabelInfo newLabelInfo = label.CreateLabelInfo();

    // Configure the label
    newLabelInfo.LabelId = "{YOUR-LABEL-GUID}";
    newLabelInfo.LabelName = "Confidential";
    newLabelInfo.AssignmentMethod = Office.Enums.MsoAssignmentMethod.PRIVILEGED;
    newLabelInfo.Justification = "Document contains sensitive information";
    newLabelInfo.IsEnabled = true;

    // Set the label
    label.SetLabel(newLabelInfo, null);

    Console.WriteLine("Sensitivity label applied successfully");
}

// Save and cleanup
workbook.Save();
workbook.Close();
excelApp.Quit();
excelApp.Dispose();
```

### Example 3: Check and Display Label (PowerPoint)

```csharp
using System;
using PowerPoint = NetOffice.PowerPointApi;
using Office = NetOffice.OfficeApi;

// Open PowerPoint presentation
PowerPoint.Application pptApp = new PowerPoint.Application();
PowerPoint.Presentation presentation = pptApp.Presentations.Open(
    @"C:\path\to\presentation.pptx",
    Microsoft.Office.Core.MsoTriState.msoFalse,
    Microsoft.Office.Core.MsoTriState.msoFalse,
    Microsoft.Office.Core.MsoTriState.msoTrue
);

// Get sensitivity label
Office.SensitivityLabel label = presentation.SensitivityLabel;
if (label != null)
{
    Office.LabelInfo labelInfo = label.GetLabel();
    if (labelInfo != null)
    {
        // Display label information
        Console.WriteLine("=== Sensitivity Label Information ===");
        Console.WriteLine($"Name: {labelInfo.LabelName}");
        Console.WriteLine($"ID: {labelInfo.LabelId}");
        Console.WriteLine($"Action ID: {labelInfo.ActionId}");
        Console.WriteLine($"Assignment: {labelInfo.AssignmentMethod}");
        Console.WriteLine($"Enabled: {labelInfo.IsEnabled}");
        Console.WriteLine($"Set Date: {labelInfo.SetDate}");
        Console.WriteLine($"Site ID: {labelInfo.SiteId}");
        Console.WriteLine($"Content Bits: {labelInfo.ContentBits}");
    }
    else
    {
        Console.WriteLine("No sensitivity label found on this presentation.");
    }
}

// Cleanup
presentation.Close();
pptApp.Quit();
pptApp.Dispose();
```

### Example 4: Complete Label Management Workflow

```csharp
using System;
using Word = NetOffice.WordApi;
using Office = NetOffice.OfficeApi;

public class SensitivityLabelManager
{
    public void ManageDocumentLabel(string documentPath)
    {
        Word.Application wordApp = null;
        Word.Document doc = null;

        try
        {
            // Initialize Word
            wordApp = new Word.Application();
            wordApp.DisplayAlerts = Word.Enums.WdAlertLevel.wdAlertsNone;

            // Open document
            doc = wordApp.Documents.Open(documentPath);

            // Get sensitivity label
            Office.SensitivityLabel label = doc.SensitivityLabel;
            if (label == null)
            {
                Console.WriteLine("SensitivityLabel API not available");
                return;
            }

            // Check existing label
            Office.LabelInfo existingLabel = label.GetLabel();
            if (existingLabel != null)
            {
                Console.WriteLine($"Current Label: {existingLabel.LabelName}");

                // Check if downgrading (requires justification)
                bool isDowngrade = CheckIfDowngrade(existingLabel.LabelId, newLabelGuid);

                if (isDowngrade)
                {
                    Console.WriteLine("Downgrade detected - justification required");
                }
            }

            // Create and apply new label
            Office.LabelInfo newLabel = label.CreateLabelInfo();
            newLabel.LabelId = "{YOUR-NEW-LABEL-GUID}";
            newLabel.LabelName = "Highly Confidential";
            newLabel.AssignmentMethod = Office.Enums.MsoAssignmentMethod.PRIVILEGED;
            newLabel.IsEnabled = true;

            // Add justification if downgrading
            if (existingLabel != null)
            {
                newLabel.Justification = "Label updated per policy requirements";
            }

            // Apply the label
            label.SetLabel(newLabel, "CustomContext");

            // Verify label was set
            Office.LabelInfo verifyLabel = label.GetLabel();
            if (verifyLabel != null && verifyLabel.LabelId == newLabel.LabelId)
            {
                Console.WriteLine("Label successfully applied and verified");
            }

            // Save document
            doc.Save();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Cleanup
            if (doc != null)
            {
                doc.Close();
                doc.Dispose();
            }

            if (wordApp != null)
            {
                wordApp.Quit();
                wordApp.Dispose();
            }
        }
    }

    private bool CheckIfDowngrade(string currentLabelId, string newLabelId)
    {
        // Implement your organization's label hierarchy logic
        // Return true if newLabel has lower sensitivity than currentLabel
        return false;
    }
}
```

## Testing Checklist

### Prerequisites
- [ ] Office 2016 or later installed
- [ ] Organization has Microsoft 365 with sensitivity labels configured
- [ ] User is signed in with Office Account
- [ ] Appropriate licensing for sensitivity labels

### Test Cases

#### Word.Document Tests
- [ ] Get SensitivityLabel property from Document
- [ ] Call GetLabel() on documents with labels
- [ ] Call GetLabel() on documents without labels
- [ ] Create LabelInfo using CreateLabelInfo()
- [ ] Set label using SetLabel()
- [ ] Verify label persists after save and reopen
- [ ] Test downgrade with justification
- [ ] Test all LabelInfo properties

#### Excel.Workbook Tests
- [ ] Get SensitivityLabel property from Workbook
- [ ] Call GetLabel() on workbooks with labels
- [ ] Call GetLabel() on workbooks without labels
- [ ] Set label and verify
- [ ] Test label inheritance with sheets
- [ ] Verify label after Save/SaveAs

#### PowerPoint.Presentation Tests
- [ ] Get SensitivityLabel property from Presentation
- [ ] Call GetLabel() on presentations with labels
- [ ] Call GetLabel() on presentations without labels
- [ ] Set label and verify
- [ ] Test with different presentation formats (.pptx, .pptm)

#### Error Handling Tests
- [ ] Test on systems without sensitivity label support
- [ ] Test with invalid Label GUIDs
- [ ] Test downgrade without justification
- [ ] Test with disabled labels
- [ ] Test without proper licensing

#### Integration Tests
- [ ] Test with SharePoint documents (SiteId property)
- [ ] Test label events (if applicable)
- [ ] Test with auto-classification
- [ ] Test content marking (ContentBits property)

## Next Steps

### 1. Build Verification (Windows Required)
```bash
# Open in Visual Studio on Windows
# Build configuration: Debug and Release
# Target: Source/NetOffice.sln
```

### 2. Unit Tests
Create unit test project for SensitivityLabel:
- Test property accessors
- Test method invocations
- Test enum values
- Mock COM interop for offline testing

### 3. Integration Tests
Create integration tests with actual Office applications:
- Requires Office 2016+ on test machine
- Requires test environment with sensitivity labels configured

### 4. Documentation
- [ ] Update CHANGELOG.md
- [ ] Add XML documentation comments (already included)
- [ ] Create user guide for SensitivityLabel usage
- [ ] Update API reference documentation

### 5. PR Finalization
- [ ] Commit all changes to `dev/441_SensitivityLabel_object_model` branch
- [ ] Push to remote repository
- [ ] Remove draft status from PR #445
- [ ] Request code review
- [ ] Address review feedback
- [ ] Merge to release branch

### 6. Release
- [ ] Include in NetOffice v1.9.8 release notes
- [ ] Announce new feature to users
- [ ] Monitor for issues and feedback

## Known Limitations

1. **Office 2016+ Only**: SensitivityLabel API is only available in Office 2016 and later
2. **Licensing Required**: Organizations need appropriate Microsoft 365 licensing
3. **Policy Dependent**: Label availability depends on organizational policy configuration
4. **Windows Only**: NetOffice build requires Windows environment
5. **Authentication**: User must be signed in with Office Account

## Implementation Notes

### Design Decisions

1. **Read-only Properties**: Document.SensitivityLabel, Workbook.SensitivityLabel, and Presentation.SensitivityLabel are implemented as read-only properties (get only), matching the VBA API design.

2. **Factory Pattern**: Uses NetOffice's `ExecuteKnownReferencePropertyGet` pattern for COM interop, ensuring proper object lifecycle management.

3. **Version Support**: All new APIs explicitly marked with `[SupportByVersion("Application", 16)]` to indicate Office 2016+ requirement.

4. **Documentation**: Comprehensive XML documentation comments included, referencing official Microsoft documentation URLs.

5. **Enum Naming**: `MsoAssignmentMethod` follows Microsoft's naming convention with uppercase enum values.

### Code Style

The implementation follows NetOffice conventions:
- Tab indentation
- XML documentation comments
- SupportByVersion attributes
- EntityType attributes
- Proper namespace organization
- Factory method usage for COM interop

## References

### Microsoft Documentation
- [Sensitivity Labels Overview](https://learn.microsoft.com/en-us/microsoft-365/compliance/sensitivity-labels)
- [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/)
- [Licensing Requirements](https://learn.microsoft.com/en-us/office365/servicedescriptions/microsoft-365-service-descriptions/microsoft-365-tenantlevel-services-licensing-guidance/microsoft-365-security-compliance-licensing-guidance)

### NetOffice Resources
- [NetOffice GitHub Repository](https://github.com/NetOfficeFw/NetOffice)
- [NetOffice Documentation](https://netoffice.io)

## Credits

Implementation created using:
- **Planning**: Claude Sonnet 4.5 (Plan mode)
- **Implementation**: Claude Haiku 4.5 (Agent mode) and Claude Sonnet 4.5
- **Date**: November 2025
- **Contributor**: @jozefizso

## Change Log

### 2025-11-05
- Created `SensitivityLabel` class
- Created `LabelInfo` class
- Created `MsoAssignmentMethod` enum
- Initial PR #445 created (draft)

### 2025-11-13
- Added `SensitivityLabel` property to `Word._Document`
- Added `SensitivityLabel` property to `Excel._Workbook`
- Added `SensitivityLabel` property to `PowerPoint._Presentation`
- Created comprehensive implementation plan document
- Implementation marked as complete, pending testing

---

**Issue**: #441
**PR**: #445
**Branch**: `dev/441_SensitivityLabel_object_model`
**Target Release**: NetOffice v1.9.8
