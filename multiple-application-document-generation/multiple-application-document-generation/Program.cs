using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Collections.Generic;
using System.IO;

namespace Mail_merge_with_.NET_objects
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the employee details as IEnumerable collection.
                    List<Template> templateData = GetTemplate();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Template", templateData);

                    document.MailMerge.RemoveEmptyParagraphs = true;
                    document.MailMerge.StartAtNewPage = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    //Creates file stream.
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        //Enable AutoDetectComplexScript property true
                        render.Settings.AutoDetectComplexScript = true;
                        //Converts Word document into PDF document
                        PdfDocument pdfDocument = render.ConvertToPDF(document);
                        pdfDocument.Save(@"../../../Output/Output.pdf");
                    }
                }
            }
        }
        /// <summary>
        /// Returns a collection of sample loan application data.
        /// </summary>
        public static List<Template> GetTemplate()
        {
            string websiteUrl = "https://www.examplebank.com";
            return new List<Template>
    {
        new Template(
            "LAN100001",
            "28-Sep-2026",
            "28-Sep-2026",
            "Anna Nagar Branch",
            new List<Applicant>
            {
                new Applicant(
                    "Mr.", "Kumarasamy",
                    "45, Nehru Road, Chennai - 600078",
                    "55, Link Road, Mumbai - 400053",
                    "1000",
                    "Plot No. 10, Anna Nagar, Chennai",
                    "Residential Property"),

                new Applicant(
                    "Mr.", "Mahesh Sharma",
                    "11, Brigade Road, Bangalore - 560025",
                    "45, Nehru Road, Chennai - 600078",
                    "1001",
                    "Commercial Complex, MG Road, Bangalore",
                    "Commercial Property")
            },
            "Vehicle",
            "25,50,000",
            "Home Loan",
            "8.5%",
            "21500",
            "05-Oct-2026",
            websiteUrl),

        new Template(
            "LAN100002",
            "28-Sep-2026",
            "28-Sep-2026",
            "MG Road Branch",
            new List<Applicant>
            {
                new Applicant(
                    "Mr.", "Kumarasamy",
                    "45, Nehru Road, Chennai - 600078",
                    "55, Link Road, Mumbai - 400053",
                    "1002",
                    "Commercial Complex, MG Road, Bangalore",
                    "Commercial Property")
            },
            "Working Capital",
            "51,00,000",
            "Business Loan",
            "7.5%",
            "68000",
            "15-Oct-2026",
            websiteUrl),

        new Template(
            "LAN100003",
            "28-Sep-2026",
            "28-Sep-2026",
            "Andheri Branch",
            new List<Applicant>
            {
                new Applicant(
                    "Mrs.", "Swetha",
                    "55, Link Road, Mumbai - 400053",
                    "45, Nehru Road, Chennai - 600078",
                    "1003",
                    "Flat 101, Sunshine Apartments, Mumbai",
                    "Apartment")
            },
            "Personal Use",
            "12,25,000",
            "Vehicle Loan",
            "6.5%",
            "25500",
            "01-Nov-2026",
            websiteUrl)
    };
        }
        /// <summary>
        /// Represents the details of an applicant.
        /// </summary>
        public class Applicant
        {
            public string applicantsTitle { get; set; }
            public string applicantsName { get; set; }
            public string applicantsCustomerAddresses { get; set; }

            public string guarantorAddress { get; set; }
            public string applicantNumber { get; set; }
            public string appCollaterlCollateral_Address { get; set; }
            public string appCollaterlnatureOfProperty { get; set; }
            public Applicant(
             string applicantsTitle,
             string applicantsName,
             string applicantsCustomerAddresses,
             string guarantorAddress,
             string applicantNumber,
             string appCollaterlCollateral_Address,
             string appCollaterlnatureOfProperty)
            {
                this.applicantsTitle = applicantsTitle;
                this.applicantsName = applicantsName;
                this.applicantsCustomerAddresses = applicantsCustomerAddresses;
                this.guarantorAddress = guarantorAddress;
                this.applicantNumber = applicantNumber;
                this.appCollaterlCollateral_Address = appCollaterlCollateral_Address;
                this.appCollaterlnatureOfProperty = appCollaterlnatureOfProperty;
            }
        }
        /// <summary>
        /// Represents a class to maintain loan application data.
        /// </summary>
        public class Template
        {
            public string lanNumber { get; set; }
            public string executionDate { get; set; }
            public string busisnessDate { get; set; }
            public string sourcingBranch { get; set; }
            public string PurposeofLoan { get; set; }
            public string totalAmount { get; set; }
            public string schemeType { get; set; }
            public string roiLabel { get; set; }
            public string appFinancialEmi { get; set; }
            public string emiStartDate { get; set; }
            public string websiteLink { get; set; }
            public List<Applicant> Applicant { get; set; }

            public Template(
             string lanNumber,
             string executionDate,
             string busisnessDate,
             string sourcingBranch,
             List<Applicant> applicant,
             string purposeofLoan,
             string totalAmount,
             string schemeType,
             string roiLabel,
             string appFinancialEmi,
             string emiStartDate,
             string websiteLink)
            {
                this.lanNumber = lanNumber;
                this.executionDate = executionDate;
                this.busisnessDate = busisnessDate;
                this.sourcingBranch = sourcingBranch;
                this.PurposeofLoan = purposeofLoan;
                this.Applicant = applicant;
                this.totalAmount = totalAmount;
                this.schemeType = schemeType;
                this.roiLabel = roiLabel;
                this.appFinancialEmi = appFinancialEmi;
                this.emiStartDate = emiStartDate;
                this.websiteLink = websiteLink;
            }
        }
    }
}