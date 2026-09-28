# dynamic-application-document-generation

## Overview
---

This sample demonstrates how to dynamically generate loan application documents using the [Syncfusion .NET Word Library (DocIO)](https://www.syncfusion.com/document-sdk/net-word-library). The application loads a Word template, performs mail merge operations with loan applicant data, and exports the generated documents to PDF format.

This is a practical example of using DocIO for automated document generation in financial services, HR, and business process automation scenarios.

## Project Description
---

The **dynamic-application-document-generation** sample showcases:

*   Loading and manipulating existing Word document templates (DOCX format)
*   Performing [mail merge operations](https://www.syncfusion.com/document-sdk/net-word-library/mail-merge) with structured data sources
*   Working with complex document templates containing multiple merge fields
*   Converting Word documents to [PDF format](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversion) programmatically
*   Handling financial and legal document generation workflows
*   Batch processing of documents with different data sets

## Sample Use Case
---

This application generates **Loan Sanction Letters** for a banking system. It demonstrates:

*   Loading a pre-designed Word template (`Template.docx`) containing merge fields
*   Creating a `Template` class to model loan applicant and collateral information
*   Populating merge fields with real applicant data (name, address, loan amount, EMI, ROI, etc.)
*   Automatically generating multiple documents from a single template
*   Converting the generated documents to PDF for distribution and archival

### Sample Data Included:

The application generates documents for three loan applicants:
1. **Rajesh Kumar** - Home Construction Loan (₹25,50,000) at 8.75% ROI
2. **Arun Sharma** - Business Expansion Loan (₹51,00,000) at 10.25% ROI
3. **Sneha Patil** - Vehicle Purchase Loan (₹12,25,000) at 9.50% ROI

## Key Technologies
---

*   **Syncfusion DocIO** - Word document processing and manipulation
*   **.NET 8.0** - Platform and runtime
*   **DocIORenderer** - Document rendering engine for PDF conversion
*   **Syncfusion PDF** - PDF document creation and manipulation

## Features Demonstrated
---

*   Loading Word documents with `FormatType.Automatic` for format detection
*   Using `MailMergeDataTable` for structured data binding
*   Configuring mail merge options (`RemoveEmptyParagraphs`, `StartAtNewPage`)
*   Executing mail merge operations with `ExecuteGroup()`
*   Converting Word documents to PDF with `DocIORenderer`
*   Enabling complex script detection (`AutoDetectComplexScript`)

## Project Structure
---

```
dynamic-application-document-generation/
├── Program.cs                 # Main application logic with mail merge and PDF conversion
├── Data/
│   └── Template.docx         # Word template with merge fields
├── Output/
│   └── Output.pdf            # Generated PDF output
└── dynamic-application-document-generation.csproj
```

## How to Use
---

1. **Prepare the Template:**
   - Ensure `Template.docx` exists in the `Data` folder
   - The template should contain merge fields matching the `Template` class properties
   - Common merge fields: `{{lanNumber}}`, `{{applicants1Name}}`, `{{sanctionedAmount}}`, etc.

2. **Run the Application:**
   ```bash
   dotnet run
   ```

3. **Output:**
   - The application will generate a merged PDF document saved as `Output/Output.pdf`
   - The PDF will contain documents for all applicants in the sample data

## Template Class Properties
---

The `Template` class contains the following merge field mappings:

**Loan Information:**
- `lanNumber` - Loan Application Number
- `versionNumber` - Document version
- `schemeType` - Type of loan (Home Loan, Business Loan, Vehicle Loan)
- `sanctionedAmount` - Sanctioned loan amount
- `ROI` - Rate of Interest
- `tenureMonths` - Loan tenure in months
- `appFinancialEmi` - Monthly EMI amount

**Applicant Information:**
- `applicants1Title`, `applicants1Name` - Primary applicant details
- `applicants1FathersName` - Applicant's father name
- `applicants1SpouseName` - Applicant's spouse name
- `applicants1CustomerAddresses` - Residential address

**Collateral Information:**
- `appCollaterl1Collateral_Address` - Property address
- `appCollaterl1BuiltupArea` - Property built-up area
- `appCollaterl1natureOfProperty` - Property type
- `appCollaterl1PINCode` - Property postal code

**Other Details:**
- `executionDate`, `busisnessDate` - Document dates
- `executionPlace`, `sourcingBranch` - Location information
- `ltvValue` - Loan-to-Value ratio
- `emiStartDate` - First EMI date
- `websiteLink` - Bank website URL

## Compatible Microsoft Word Versions
---

*   Microsoft Word 97-2003
*   Microsoft Word 2007
*   Microsoft Word 2010
*   Microsoft Word 2013
*   Microsoft Word 2016
*   Microsoft Word 2019
*   Microsoft 365

## Supported File Formats
---

*   Creates, reads, and edits popular text file formats like [DOC](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#doc-to-docx-and-docx-to-doc), DOT, [DOCM](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#macros-docm-dotm), DOTM, [DOCX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-document-docx), [DOTX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-template-dotx), [HTML](https://www.syncfusion.com/document-sdk/net-word-library/html-conversions), [RTF](https://help.syncfusion.com/document-processing/word/conversions/rtf-conversions), [TXT](https://www.syncfusion.com/document-sdk/net-word-library/text-conversions), [Markdown](https://help.syncfusion.com/document-processing/word/conversions/markdown-to-word-conversion)and [XML (WordML)](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-processing-xml-xml).
*   Converts Word documents also to [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversionhub-docio-examples), [Image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion), and [ODT](https://help.syncfusion.com/document-processing/word/conversions/word-to-odt-conversion) files.

## How to run the examples
---

*   Download this project to a location in your disk.
*   Open the solution file using Visual Studio.
*   Rebuild the solution to install the required NuGet packages.
*   Run the application.

## Resources
---

*   **Product page:** [Syncfusion® Word Framework](https://www.syncfusion.com/document-sdk/net-word-library)
*   **Documentation:** [Syncfusion® Word library](https://help.syncfusion.com/document-processing/word/word-library/net/overview)
*   **Online demo:** [Syncfusion® Word library - Online demos](https://document.syncfusion.com/demos/word/salesinvoice#/bootstrap5)
*   **Blog:** [Syncfusion® Word library - Blog](https://www.syncfusion.com/blogs/category/docio?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)
*   **Knowledge Base:** [Syncfusion® Word library - Knowledge Base](https://www.syncfusion.com/kb/aspnetcore/docio?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)
*   **Ebooks:** [Syncfusion® Word library - Ebooks](https://www.syncfusion.com/succinctly-free-ebooks?utm_source=nuget&utm_medium=listing&utm_campaign=aspnetcore-docio-nuget)
*   **FAQ:** [Syncfusion® Word library - FAQ](https://www.syncfusion.com/faq/?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)

## Support and feedback
---

*   For any other queries, reach our [Syncfusion® support team](https://www.syncfusion.com/support/directtrac/incidents/newincident?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).
*   Request new feature through [Syncfusion® feedback portal](https://www.syncfusion.com/feedback?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

## License
---

This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/eula/es/?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples). You can purchase a licnense [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

## About Syncfusion®
---

Founded in 2001 and headquartered in Research Triangle Park, N.C., Syncfusion® has more than 29,000 customers and more than 1 million users, including large financial institutions, Fortune 500 companies, and global IT consultancies.

Today, we provide 1700+ components and frameworks for web ([Blazor](https://www.syncfusion.com/blazor-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [ASP.NET Core](https://www.syncfusion.com/aspnet-core-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [ASP.NET MVC](https://www.syncfusion.com/aspnet-mvc-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [ASP.NET WebForms](https://www.syncfusion.com/jquery/aspnet-webforms-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [JavaScript](https://www.syncfusion.com/javascript-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Angular](https://www.syncfusion.com/angular-ui-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [React](https://www.syncfusion.com/react-ui-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Vue](https://www.syncfusion.com/vue-ui-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), and [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)), mobile ([Xamarin](https://www.syncfusion.com/xamarin-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [UWP](https://www.syncfusion.com/uwp-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), and [JavaScript](https://www.syncfusion.com/javascript-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [.NET MAUI](https://www.syncfusion.com/maui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)) and desktop development ([WinForms](https://www.syncfusion.com/winforms-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [WPF](https://www.syncfusion.com/wpf-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [WinUI](https://www.syncfusion.com/winui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [UWP](https://www.syncfusion.com/uwp-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), and [.NET MAUI](https://www.syncfusion.com/maui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)) a. We provide ready-to-deploy enterprise software for dashboards, reports, data integration, and big data processing. Many customers have saved millions in licensing fees by deploying our software.