using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System;
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
                    MailMergeDataTable dataSource = new MailMergeDataTable("Template", templateData);
                    document.MailMerge.RemoveEmptyParagraphs = true;
                    document.MailMerge.StartAtNewPage = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);
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
        /// Gets the employee details to perform mail merge.
        /// </summary>
        public static List<Template> GetTemplate()
        {
            List<Template> template = new List<Template>();

            template.Add(new Template(
                "LAN100001",
                "V1.0",
                "28-Sep-2026",
                "28-Sep-2026",
                "Chennai",
                "Anna Nagar Branch",
                "Mr.",
                "Rajesh Kumar",
                "Kumarasamy",
                "Lakshmi Rajesh",
                "12, Gandhi Street, Chennai - 600040",
                "45, Nehru Road, Chennai - 600078",
                "Home Construction",
                "",
                "2500000",
                "50000",
                "2550000",
                "Twenty Five Lakh Fifty Thousand Only",
                "240",
                "Home Loan",
                "+",
                "8.75%",
                "Floating ROI",
                "21500",
                "Twenty One Thousand Five Hundred Only",
                "05-Oct-2026",
                "https://www.examplebank.com",
                "75%",
                "Plot No. 10, Anna Nagar, Chennai",
                "600040",
                "1800 Sq.ft",
                "Residential Property",
                "Mrs.",
                "Priya Rajesh"));

            template.Add(new Template(
                "LAN100002",
                "V1.1",
                "28-Sep-2026",
                "28-Sep-2026",
                "Bangalore",
                "MG Road Branch",
                "Mr.",
                "Arun Sharma",
                "Mahesh Sharma",
                "Pooja Arun",
                "22, Residency Road, Bangalore - 560001",
                "11, Brigade Road, Bangalore - 560025",
                "Business Expansion",
                "Working Capital",
                "5000000",
                "100000",
                "5100000",
                "Fifty One Lakh Only",
                "120",
                "Business Loan",
                "+",
                "10.25%",
                "Fixed ROI",
                "68000",
                "Sixty Eight Thousand Only",
                "15-Oct-2026",
                "https://www.examplebank.com",
                "68%",
                "Commercial Complex, MG Road, Bangalore",
                "560001",
                "3200 Sq.ft",
                "Commercial Property",
                "Mrs.",
                "Pooja Sharma"));

            template.Add(new Template(
                "LAN100003",
                "V2.0",
                "28-Sep-2026",
                "28-Sep-2026",
                "Mumbai",
                "Andheri Branch",
                "Ms.",
                "Sneha Patil",
                "Ramesh Patil",
                "N/A",
                "101, Sunshine Apartments, Andheri East, Mumbai",
                "55, Link Road, Mumbai - 400053",
                "Vehicle Purchase",
                "Personal Use",
                "1200000",
                "25000",
                "1225000",
                "Twelve Lakh Twenty Five Thousand Only",
                "60",
                "Vehicle Loan",
                "+",
                "9.50%",
                "Vehicle ROI",
                "25500",
                "Twenty Five Thousand Five Hundred Only",
                "01-Nov-2026",
                "https://www.examplebank.com",
                "85%",
                "Flat 101, Sunshine Apartments, Mumbai",
                "400059",
                "1100 Sq.ft",
                "Apartment",
               "Mrs.",
                "Pooja"));

            return template;
        }
    }

    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Template
    {
        public string lanNumber { get; set; }
        public string versionNumber { get; set; }
        public string executionDate { get; set; }
        public string busisnessDate { get; set; }
        public string executionPlace { get; set; }
        public string sourcingBranch { get; set; }
        public string applicants1Title { get; set; }
        public string applicants1Name { get; set; }

        public string applicants1FathersName { get; set; }
        public string applicants1SpouseName { get; set; }
        public string applicants1CustomerAddresses { get; set; }

        public string guarantor1Address { get; set; }
        public string PurposeofLoan { get; set; }
        public string EndUse { get; set; }

        public string sanctionedAmount { get; set; }
        public string insuranceAmount { get; set; }
        public string totalAmount { get; set; }

        public string totalAmountInWords { get; set; }
        public string tenureMonths { get; set; }
        public string schemeType { get; set; }

        public string addSign { get; set; }
        public string ROI { get; set; }
        public string roiLabel { get; set; }

        public string appFinancialEmi { get; set; }
        public string emiInWord { get; set; }

        public string emiStartDate { get; set; }

        public string websiteLink { get; set; }
        public string ltvValue { get; set; }

        public string appCollaterl1Collateral_Address { get; set; }

        public string appCollaterl1PINCode { get; set; }
        public string appCollaterl1BuiltupArea { get; set; }

        public string appCollaterl1natureOfProperty { get; set; }

        public string applicants2Title { get; set; }
        public string applicants2Name { get; set; }
        public Template(
         string lanNumber,
         string versionNumber,
         string executionDate,
         string busisnessDate,
         string executionPlace,
         string sourcingBranch,
         string applicants1Title,
         string applicants1Name,
         string applicants1FathersName,
         string applicants1SpouseName,
         string applicants1CustomerAddresses,
         string guarantor1Address,
         string purposeofLoan,
         string endUse,
         string sanctionedAmount,
         string insuranceAmount,
         string totalAmount,
         string totalAmountInWords,
         string tenureMonths,
         string schemeType,
         string addSign,
         string roi,
         string roiLabel,
         string appFinancialEmi,
         string emiInWord,
         string emiStartDate,
         string websiteLink,
         string ltvValue,
         string appCollaterl1Collateral_Address,
         string appCollaterl1PINCode,
         string appCollaterl1BuiltupArea,
         string appCollaterl1natureOfProperty,
         string applicants2Title,
         string applicants2Name)
        {
            this.lanNumber = lanNumber;
            this.versionNumber = versionNumber;
            this.executionDate = executionDate;
            this.busisnessDate = busisnessDate;
            this.executionPlace = executionPlace;
            this.sourcingBranch = sourcingBranch;
            this.applicants1Title = applicants1Title;
            this.applicants1Name = applicants1Name;
            this.applicants1FathersName = applicants1FathersName;
            this.applicants1SpouseName = applicants1SpouseName;
            this.applicants1CustomerAddresses = applicants1CustomerAddresses;
            this.guarantor1Address = guarantor1Address;
            this.PurposeofLoan = purposeofLoan;
            this.EndUse = endUse;
            this.sanctionedAmount = sanctionedAmount;
            this.insuranceAmount = insuranceAmount;
            this.totalAmount = totalAmount;
            this.totalAmountInWords = totalAmountInWords;
            this.tenureMonths = tenureMonths;
            this.schemeType = schemeType;
            this.addSign = addSign;
            this.ROI = roi;
            this.roiLabel = roiLabel;
            this.appFinancialEmi = appFinancialEmi;
            this.emiInWord = emiInWord;
            this.emiStartDate = emiStartDate;
            this.websiteLink = websiteLink;
            this.ltvValue = ltvValue;
            this.appCollaterl1Collateral_Address = appCollaterl1Collateral_Address;
            this.appCollaterl1PINCode = appCollaterl1PINCode;
            this.appCollaterl1BuiltupArea = appCollaterl1BuiltupArea;
            this.appCollaterl1natureOfProperty = appCollaterl1natureOfProperty;
            this.applicants2Title = applicants2Title;
            this.applicants2Name = applicants2Name;
        }
    }
}
