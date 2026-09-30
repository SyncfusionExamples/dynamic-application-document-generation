using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Collections.Generic;
using System.IO;

namespace dynamic_application_document_generation
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
                    List<Applicant> applicantList = GetApplicantData();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("Applicant", applicantList);
                    //Performs Group Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);
                    // Define the merge field names available in the template.
                    string[] fieldNames = new string[] { "lanNumber", "executionDate", "busisnessDate", "sourcingBranch",
                                                         "PurposeofLoan","totalAmount","schemeType","roiLabel",
                                                           "appFinancialEmi","emiStartDate","websiteLink"};

                    // Define the corresponding values to populate the merge fields.
                    string[] fieldValues = new string[]
                                            {
                                            "LAN-1001", // lanNumber
                                            "30/09/2026", // executionDate
                                            "30/09/2026", // busisnessDate
                                            "Central Branch", // sourcingBranch
                                            "Home Loan", // PurposeofLoan
                                            "50,50,000", // totalAmount
                                            "Floating Rate", // schemeType
                                            "8.5%", // roiLabel
                                            "43,400", // appFinancialEmi
                                            "01/11/2026", // emiStartDate
                                            "https://www.examplebank.com", // websiteLink
                                            };
                    
                    // Execute a simple mail merge for single-value fields.
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    // Retrieve applicant details for group mail merge.
                    //Creates file stream.
                    using (DocIORenderer render = new DocIORenderer())
                    {
                        // Automatically detect and render complex scripts properly.
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
        public static List<Applicant> GetApplicantData()
        {
            List<Applicant> applicantData = new List<Applicant>
                                    {
                                    new Applicant("1",
                                    "Mr.",
                                    "John Smith",
                                    "Robert Smith",
                                    "Mary Smith",
                                    "12 Baker Street, London",
                                    "45 King Street, London",
                                    "Flat No. 101, Green Residency, Baker Street, London", 
                                    "Residential"
                                    ),

                                    new Applicant("2",
                                    "Mrs.",
                                    "Jennifer Brown",
                                    "David Brown",
                                    "Michael Brown",
                                    "78 Oxford Road, Manchester",
                                    "22 Queen Avenue, Manchester",
                                    "Villa No. 25, Palm Gardens, Oxford Road, Manchester",
                                    "Residential"
                                    ),

                                    new Applicant("3",
                                    "Mr.",
                                    "William Johnson",
                                    "Thomas Johnson",
                                    "Sophia Johnson",
                                    "33 Victoria Street, Birmingham",
                                    "99 Albert Road, Birmingham",
                                    "Plot No. 56, Lake View Layout, Victoria Street, Birmingham",
                                    "Commercial"
                                    )
                                    };

            return applicantData;
        }
    }
    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Applicant
    {
        public string applicantsTitle { get; set; }
        public string applicantsName { get; set; }

        public string applicantsFathersName { get; set; }
        public string applicantsSpouseName { get; set; }
        public string applicantsCustomerAddresses { get; set; }

        public string guarantorAddress { get; set; }
        public string ApplicantNumber { get; set; }
        public string appCollaterlCollateral_Address { get; set; }
        public string appCollaterlnatureOfProperty { get; set; }

        public Applicant(
         string applicationNumber,
         string applicantsTitle,
         string applicantsName,
         string applicantsFathersName,
         string applicantsSpouseName,
         string applicantsCustomerAddresses,
         string guarantorAddress,
         string appCollaterlCollateral_Address,
         string appCollaterlnatureOfProperty
         )
        {
            this.ApplicantNumber = applicationNumber;
            this.applicantsTitle = applicantsTitle;
            this.applicantsName = applicantsName;
            this.applicantsFathersName = applicantsFathersName;
            this.applicantsSpouseName = applicantsSpouseName;
            this.applicantsCustomerAddresses = applicantsCustomerAddresses;
            this.guarantorAddress = guarantorAddress;
            this.appCollaterlCollateral_Address = appCollaterlCollateral_Address;
            this.appCollaterlnatureOfProperty = appCollaterlnatureOfProperty;
        }
    }
}
