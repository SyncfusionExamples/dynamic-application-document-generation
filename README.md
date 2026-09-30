# Dynamic Application Document Generation

## Overview
---

This repository contains two sample applications that demonstrate how to dynamically generate loan application documents using the [Syncfusion .NET Word Library (DocIO)](https://www.syncfusion.com/document-sdk/net-word-library). Both samples load a Word template, populate merge fields with banking data, and export the generated document to PDF format.

The first sample focuses on a single application template with grouped applicant records, while the second sample demonstrates nested mail merge for multiple loan application records, each containing multiple applicants. Together, they show practical ways to automate financial document generation for banking, lending, and business process workflows.

## Project Description
---

The repository includes two sample projects:

*   **dynamic-application-document-generation** - demonstrates a grouped mail merge approach for a single document template with multiple applicant records
*   **multiple-application-document-generation** - demonstrates nested group mail merge for multiple loan applications, each containing applicant details and collateral information

These samples showcase:

*   Loading and manipulating existing Word document templates (DOCX format)
*   Performing [mail merge operations](https://www.syncfusion.com/document-sdk/net-word-library/mail-merge) with strongly typed data objects and nested collections
*   Working with complex document templates containing multiple merge fields
*   Converting Word documents to [PDF format](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversion) programmatically
*   Handling financial document generation workflows for loan sanction letters and application forms
*   Generating multiple records from a single template using grouped data structures

## Sample Use Case
---

These applications generate **loan application and sanction documents** for a banking system. They demonstrate:

*   Loading a pre-designed Word template (`Template.docx`) containing merge fields
*   Creating a `Template` class to model loan application and applicant details
*   Populating merge fields with real data such as loan reference number, branch name, purpose, amount, EMI, ROI, and website URL
*   Automatically generating multiple documents from a single template
*   Converting the merged documents to PDF for distribution, review, and archival

### Sample Data Included:

The sample data includes a range of loan scenarios such as:
1. **Vehicle Loan** - multiple applicant data with residential and commercial property details
2. **Home Loan** - loan structure with EMI and property information
3. **Business Loan** - working capital application with nested applicant records

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
*   Configuring mail merge options such as `RemoveEmptyParagraphs` and `StartAtNewPage`
*   Executing grouped and nested mail merge operations with `ExecuteGroup()` and `ExecuteNestedGroup()`
*   Converting Word documents to PDF with `DocIORenderer`
*   Enabling complex script detection with `AutoDetectComplexScript`
*   Mapping nested object collections to document fields for multi-application generation

## Project Structure
---

```
dynamic-application-document-generation/
├── dynamic-application-document-generation.sln
├── dynamic-application-document-generation/
│   ├── Program.cs                 # Group mail merge and PDF conversion sample
│   ├── Data/
│   │   └── Template.docx         # Word template with merge fields
│   ├── Output/
│   │   └── Output.pdf            # Generated PDF output
│   └── dynamic-application-document-generation.csproj
│
├── multiple-application-document-generation/
│   ├── multiple-application-document-generation.sln
│   ├── multiple-application-document-generation/
│   │   ├── Program.cs            # Nested group mail merge and PDF conversion sample
│   │   ├── Data/
│   │   │   └── Template.docx    # Word template with nested merge fields
│   │   ├── Output/
│   │   │   └── Output.pdf       # Generated PDF output
│   │   └── multiple-application-document-generation.csproj
│   └── Output/
└── README.md
```

## How to Use
---

1. **Open the solution:**
   - Open either `dynamic-application-document-generation.sln` or `multiple-application-document-generation.sln` using Visual Studio.

2. **Prepare the Template:**
   - Ensure `Template.docx` exists in the `Data` folder.
   - The template should contain merge fields matching the object properties used in the sample.

3. **Run the Application:**
   ```bash
   dotnet run
   ```

4. **Output:**
   - Each project generates a merged PDF saved in its `Output` folder.
   - The generated PDF contains the loan application data populated from the sample records.

## Template Class Properties
---

The sample data model contains the following merge field mappings.

**Loan Information:**
- `lanNumber` - Loan application number
- `executionDate` - Document execution date
- `busisnessDate` - Business date for processing
- `sourcingBranch` - Branch or processing location
- `PurposeofLoan` - Loan purpose
- `totalAmount` - Sanctioned loan amount
- `schemeType` - Loan scheme or rate plan
- `roiLabel` - Rate of interest
- `appFinancialEmi` - Monthly EMI amount
- `emiStartDate` - EMI start date
- `websiteLink` - Bank website URL

**Applicant Information:**
- `applicantsTitle` - Applicant title
- `applicantsName` - Applicant name
- `applicantsFathersName` - Applicant father's name
- `applicantsSpouseName` - Applicant spouse name
- `applicantsCustomerAddresses` - Applicant address
- `guarantorAddress` - Guarantor or reference address

**Collateral Information:**
- `appCollaterlCollateral_Address` - Property or collateral address
- `appCollaterlnatureOfProperty` - Property type
- `applicantNumber` - Applicant reference number

**Nested Collection Mapping:**
- `Applicant` - Collection of applicants nested under each loan application record

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

*   Creates, reads, and edits popular text file formats like [DOC](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#doc-to-docx-and-docx-to-doc), DOT, [DOCM](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#macros-docm-dotm), DOTM, [DOCX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-document-docx), [DOTX](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-template-dotx), [HTML](https://www.syncfusion.com/document-sdk/net-word-library/html-conversions), [RTF](https://help.syncfusion.com/document-processing/word/conversions/rtf-conversions), [TXT](https://www.syncfusion.com/document-sdk/net-word-library/text-conversions), [Markdown](https://help.syncfusion.com/document-processing/word/conversions/markdown-to-word-conversion) and [XML (WordML)](https://help.syncfusion.com/document-processing/word/conversions/word-file-formats-conversions#word-processing-xml-xml).
*   Converts Word documents also to [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversionhub-docio-examples), [Image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion), and [ODT](https://help.syncfusion.com/document-processing/word/conversions/word-to-odt-conversion) files.

## How to run the examples
---

*   Download this project to a location on your disk.
*   Open the relevant solution file using Visual Studio.
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

This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/eula/es/?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples). You can purchase a license [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples).

## About Syncfusion®
---

Founded in 2001 and headquartered in Research Triangle Park, N.C., Syncfusion® has more than 29,000 customers and more than 1 million users, including large financial institutions, Fortune 500 companies, and global IT consultancies.

Today, we provide 1700+ components and frameworks for web ([Blazor](https://www.syncfusion.com/blazor-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [ASP.NET Core](https://www.syncfusion.com/aspnet-core-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [ASP.NET MVC](https://www.syncfusion.com/aspnet-mvc-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [ASP.NET WebForms](https://www.syncfusion.com/jquery/aspnet-webforms-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [JavaScript](https://www.syncfusion.com/javascript-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Angular](https://www.syncfusion.com/angular-ui-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [React](https://www.syncfusion.com/react-ui-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Vue](https://www.syncfusion.com/vue-ui-components?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), and [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)), mobile ([Xamarin](https://www.syncfusion.com/xamarin-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [UWP](https://www.syncfusion.com/uwp-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), and [JavaScript](https://www.syncfusion.com/javascript-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [.NET MAUI](https://www.syncfusion.com/maui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)) and desktop development ([WinForms](https://www.syncfusion.com/winforms-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [WPF](https://www.syncfusion.com/wpf-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [WinUI](https://www.syncfusion.com/winui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [Flutter](https://www.syncfusion.com/flutter-widgets?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), [UWP](https://www.syncfusion.com/uwp-ui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples), and [.NET MAUI](https://www.syncfusion.com/maui-controls?utm_source=github&utm_medium=listing&utm_campaign=github-docio-examples)) and provide ready-to-deploy enterprise software for dashboards, reports, data integration, and big data processing. Many customers have saved millions in licensing fees by deploying our software.
