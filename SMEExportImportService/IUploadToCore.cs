using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace SMEExportImportService
{
    [ServiceContract]
    public interface IUploadToCore
    {
        [OperationContract]
        [FaultContract(typeof(ServicesData))]
        string CreateUploadFile(string regno, string type, string prk);
    }

    [DataContract]
    public class ServicesData
    {
        [DataMember]
        public bool Result { get; set; }
        [DataMember]
        public string ErrorMessage { get; set; }
        [DataMember]
        public string ErrorDetails { get; set; }
    }
}
