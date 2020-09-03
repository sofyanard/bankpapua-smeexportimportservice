using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Runtime.Serialization;

namespace SMEExportImportService
{
    [ServiceContract]
    public interface IWord
    {
        [OperationContract]
        void ExportWord(string name);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string DocumentExportASCXCreateWord(string templateid, string regno, string userid);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string DocumentExportASCXCreateWordPk(string templateid, string regno, string seq, string userid);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string DocumentExportASCXCreateExcel(string templateid, string regno, string userid);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string DocumentExportASCXCreateExcel2(string templateid, string regno, string userid);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string DocumentUploadASCXReadExcel(string filename, string templateid, string regno);

        [OperationContract]
        [FaultContract(typeof (ServiceData))]
        string DocumentUploadASCXReadExcelPensiun(string filename, string templateid, string regno);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string Neraca_KMK_KI_SMALLASPXviewExcel(string dir, string regno, string userid, out Dictionary<string, string> results);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string Neraca_KMK_KI_SMALLASPXviewExcel_LabaRugi(string dir, string regno, string userid, out Dictionary<string, string> results);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string Neraca_KMK_KI_MediumASPXViewExcel(string dir, string regno, string userid, out Dictionary<string, string> results);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string Neraca_KMK_KI_MediumASPXviewExcel_LabaRugi(string directori, string regno, string userid, out Dictionary<string, string> results);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string AppraisalNewASPXReadExcel(string filename, string templateid, string regno, string curef, string clseq);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainExport_Word(string regno, string userid, string var_idExport1, string var_idExport2);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateNotaWord(string regno, string userid, string sessionfullName, string branchName, string ddl_manualSelectedValue, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateNotaExcel(string regno, string userid, string SessionBranchName, string SessionFullName, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateNew(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateExist(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateSyarat(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateSyarat2(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateRata(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateBank(string regno, string userid, string SessionFullName, string SessionBranchName, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CreditProposalMainp_CreateUrus(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string SPPKExportASPXCreateSPPKWord(string regno, string userid, out Dictionary<string, string> result);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CustomerInfoExportASPXExport_Excel(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CustomerInfoExportASPXExport_Word(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CBICustomerInfoExportASPXExport_Excel(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue);
        [OperationContract]
        [FaultContract(typeof(ServiceData))]
        string CBICustomerInfoExportASPXExport_Word(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue);
    }

    [DataContract]
    public class ServiceData
    {
        [DataMember]
        public bool Result { get; set; }
        [DataMember]
        public string ErrorMessage { get; set; }
        [DataMember]
        public string ErrorDetails { get; set; }
    }
}
