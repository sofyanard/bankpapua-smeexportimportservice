using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Configuration;
using System.IO;
using DMS.DBConnection;
using System.Data;

namespace SMEExportImportService
{
    class ScanningAlphabitFolder
    {
        private static Connection conn = null;

        public static void Scanning(object obj, DataTable dtABFile, DataTable dtParentLooper, DataTable dtABStatus, DataTable dtABStatusDetail)
        {
            string path = (string)obj;

            foreach (string file in Directory.GetFiles(path))
            {
                //
                try
                {
                    string line = "";
                    string[] stringSeparators = new string[] { "\r\n" };
                    string[] result;
                    string preNameFile = "";

                    /*
                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                    conn.QueryString = "INSERT INTO AB_TEST VALUES('" + file + "')";
                    conn.ExecuteQuery();
                    conn.CloseConnection();
                    */

                    using (StreamReader reader = new StreamReader(file))
                    {
                        line = reader.ReadToEnd();
                        result = line.Split(stringSeparators, StringSplitOptions.None);
                        reader.Close();

                        if (file.Contains("CIF"))
                        {
                            preNameFile = "CIF";
                        }
                        else if (file.Contains("COL"))
                        {
                            preNameFile = "COL";
                        }
                        DataRow[] dataABFile = dtABFile.Select("AB_FILE_NAME = '" + preNameFile + "'");
                        string ID_AB_FILE = dataABFile[0]["ID_AB_FILE"].ToString();

                        //little hard code wont hurt you
                        string regno = result[1].Length > 17 ? result[1].Substring(result[1].Length - 17, 17) : result[1];

                        while (true)
                        {
                            conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                            conn.QueryString = "SELECT * FROM APPLICATION WHERE AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();

                            if (conn.GetRowCount() == 0)
                            {
                                regno = regno.Remove(0, 1);
                            }
                            else
                            {
                                break;
                            }
                            conn.CloseConnection();
                        }

                        //kosongin error table
                        conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                        conn.QueryString = "DELETE AB_ERROR_UPLOAD WHERE AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();
                        conn.CloseConnection();
                        //------------------------------------------------------------------------------

                        int CURRENT_LINE = 0;
                        int TOTAL_LINE = 0;

                        DataRow[] dataParentLooper = dtParentLooper.Select("ID_AB_FILE = '" + ID_AB_FILE + "'");
                        bool error = false;
                        for (int i = 0; i < result.Count(); i++)
                        {
                            for (int j = 0; j < dataParentLooper.Length; j++)
                            {
                                //harus tau looping apa g
                                //harus tau baris keberapa dimasukin kemana

                                string ID_PARENT = dataParentLooper[j]["ID_PARENT"].ToString();
                                string COUNTING = dataParentLooper[j]["COUNTING"].ToString();
                                string READ_FEEDBACK = dataParentLooper[j]["READ_FEEDBACK"].ToString();

                                int rowCount = 0;
                                if (COUNTING == "")
                                {
                                    rowCount = 1;
                                }
                                else
                                {
                                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                    conn.QueryString = COUNTING.Replace("#AP_REGNO", "'" + regno + "'");
                                    conn.ExecuteQuery();
                                    rowCount = conn.GetRowCount();
                                    conn.CloseConnection();
                                }

                                DataRow[] ABStatusRow = dtABStatus.Select("ID_PARENT = " + ID_PARENT);
                                if (ABStatusRow.Length == 0)
                                {
                                    continue;
                                }

                                for (int k = 0; k < rowCount; k++)
                                {
                                    string ABStatusRow_DESKRIPSI = "";
                                    string ABStatusRow_LINENUM = "";
                                    string ABStatusRow_LOOPINGCOUNTER = "";

                                    for (int l = 0; l < ABStatusRow.Length; l++)
                                    {
                                        ABStatusRow_DESKRIPSI = ABStatusRow[l]["DESCRIPT"].ToString();
                                        ABStatusRow_LINENUM = ABStatusRow[l]["LINENUM"].ToString();
                                        ABStatusRow_LOOPINGCOUNTER = ABStatusRow[l]["LOOPINGCOUNTER"].ToString();

                                        //baca 
                                        //kalo hasilnya fail baris mana yg fail infokan
                                        //kalo hasilnya sukses baris mana yg sukses, update bookedprod
                                        CURRENT_LINE = TOTAL_LINE + (int.Parse(ABStatusRow_LINENUM) - 1);
                                        string status = result[CURRENT_LINE];

                                        if (status == "C")
                                        {
                                            conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                            conn.QueryString = "SELECT [DESC] as DESCRIPT, LINENUM, QUERY_MARKING, ID_AB_STATUS_DETAIL, ID_AB_PARENT FROM AB_STATUS_DETAIL WHERE ID_AB_PARENT = '" + ID_PARENT + "' ORDER BY ID_AB_STATUS_DETAIL";
                                            conn.ExecuteQuery();
                                            conn.CloseConnection();

                                            string PRODUCTID = "";
                                            string APPTYPE = "";
                                            string PROD_SEQ = "";
                                            string CIF = "";
                                            string ACC_SEQ = "";
                                            string LIMITDISBURSED = "";
                                            string ACC_NO = "";

                                            DataTable ABStatusDetailDT = conn.GetDataTable();
                                            for (int m = 0; m < ABStatusDetailDT.Rows.Count; m++)
                                            {
                                                string DESKRIPSI = ABStatusDetailDT.Rows[m]["DESCRIPT"].ToString();
                                                string LINENUM = ABStatusDetailDT.Rows[m]["LINENUM"].ToString();
                                                string QUERY_MARKING = ABStatusDetailDT.Rows[m]["QUERY_MARKING"].ToString();
                                                string ID_AB_STATUS_DETAIL = ABStatusDetailDT.Rows[m]["ID_AB_STATUS_DETAIL"].ToString();

                                                int post = 0;
                                                post = TOTAL_LINE + (int.Parse(LINENUM) - 1);
                                                string key = result[post];

                                                if (DESKRIPSI == "Regno")
                                                {
                                                    key = regno;
                                                }

                                                if (READ_FEEDBACK != "")
                                                {
                                                    if (ID_AB_STATUS_DETAIL == "7" || ID_AB_STATUS_DETAIL == "4")
                                                    {
                                                        PRODUCTID = key;
                                                    }
                                                    else if (ID_AB_STATUS_DETAIL == "8")
                                                    {
                                                        ACC_SEQ = key;
                                                    }
                                                    else if (ID_AB_STATUS_DETAIL == "5")
                                                    {
                                                        conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                                        conn.QueryString = "SELECT PRODUCTID FROM RFPRODUCT WHERE SIBS_PRODCODE = '" + key + "'";
                                                        conn.ExecuteQuery();
                                                        key = conn.GetFieldValue("PRODUCTID");
                                                        conn.CloseConnection();
                                                    }
                                                    READ_FEEDBACK = READ_FEEDBACK.Replace(QUERY_MARKING, key);
                                                }
                                            }

                                          
                                            if (READ_FEEDBACK != "")
                                            {
                                                if (ID_PARENT == "3")
                                                {
                                                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                                    conn.QueryString = "SELECT APPTYPE, CU_CIF, CP_LIMIT, PROD_SEQ, CUSTPRODUCT.PRODUCTID, CUSTOMER.CU_REF, RFPRODUCT.IS_PRK FROM CUSTOMER, APPLICATION, CUSTPRODUCT, RFPRODUCT WHERE CUSTOMER.CU_REF = APPLICATION.CU_REF AND APPLICATION.AP_REGNO = CUSTPRODUCT.AP_REGNO AND CUSTPRODUCT.AP_REGNO = '" + regno + "' AND CUSTPRODUCT.PRODUCTID = RFPRODUCT.PRODUCTID AND RFPRODUCT.SIBS_PRODCODE = '" + PRODUCTID + "'";
                                                    conn.ExecuteQuery();
                                                    APPTYPE = conn.GetFieldValue(0, 0);
                                                    CIF = conn.GetFieldValue(0, 1);
                                                    LIMITDISBURSED = conn.GetFieldValue(0, 2);
                                                    PROD_SEQ = conn.GetFieldValue(0, 3);
                                                    string PRODUCTIDLOS = conn.GetFieldValue(0, 4);
                                                    string CUREF = conn.GetFieldValue(0, 5);
                                                    string PRK = conn.GetFieldValue(0, 6);

                                                    //simpan disini nomor rekening
                                                    ACC_NO = result[87];
                                                    if (PRK == "1")
                                                    {
                                                        conn.QueryString = "UPDATE COMPANY_INFO SET CI_BMSAVINGACCNO_PRK = '" + ACC_NO + "' WHERE CU_REF = '" + CUREF + "'";
                                                        conn.ExecuteQuery();
                                                    }
                                                    else
                                                    {
                                                        conn.QueryString = "UPDATE COMPANY_INFO SET CI_BMSAVINGACCNO = '" + ACC_NO + "' WHERE CU_REF = '" + CUREF + "'";
                                                        conn.ExecuteQuery();
                                                    }

                                                    conn.QueryString = "exec uploadtext_file_bookcustprod '" + regno + "', '" + APPTYPE + "', '" + PRODUCTIDLOS + "', '" + PROD_SEQ + "', '" + CIF + "', '" + ACC_SEQ + "', '" + LIMITDISBURSED + "', '" + ACC_NO + "'";
                                                    conn.ExecuteQuery();

                                                    if (preNameFile == "CIF")
                                                    {
                                                        /*
                                                             @AP_REGNO
                                                            @PRODUCTID
                                                            @APPTYPE
                                                            @USERID
                                                            @PROD_SEQ
                                                            @PG_TRACK
                                                         */

                                                        conn.QueryString = "UPDATE CUSTPRODUCT SET CP_NOTES = 'LOAN DONE' WHERE AP_REGNO = '" + regno + "' AND PRODUCTID = '" + PRODUCTIDLOS + "'";
                                                        conn.ExecuteQuery();

                                                        conn.QueryString = "UPDATE CUSTPRODUCT SET ACC_SEQ = '" + ACC_SEQ + "' WHERE AP_REGNO = '" + regno + "' AND PROD_SEQ = '" + PROD_SEQ + "' AND PRODUCTID = '" + PRODUCTIDLOS + "' AND APPTYPE = '" + APPTYPE + "'";
                                                        conn.ExecuteQuery();

                                                        conn.QueryString = "exec TRACKUPDATE '" + regno + "', '" + PRODUCTIDLOS + "','" + APPTYPE + "','System','" + PROD_SEQ + "','BP17.0'";
                                                        conn.ExecuteQuery();
                                                    }
                                                    conn.CloseConnection();
                                                }
                                                else
                                                {
                                                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                                    conn.QueryString = READ_FEEDBACK;
                                                    conn.ExecuteQuery();
                                                    conn.CloseConnection();
                                                }
                                            }
                                        }
                                        else if (status == "F")
                                        {
                                            conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                            conn.QueryString = "SELECT [DESC] as DESCRIPT, LINENUM, QUERY_MARKING, ID_AB_STATUS_DETAIL, ID_AB_PARENT FROM AB_STATUS_DETAIL WHERE ID_AB_PARENT = '" + ID_PARENT + "' ORDER BY ID_AB_STATUS_DETAIL";
                                            conn.ExecuteQuery();

                                            string infoTambahan = "";

                                            string PRODUCTID = "";
                                            string APPTYPE = "";
                                            string PROD_SEQ = "";
                                            string CIF = "";
                                            string ACC_SEQ = "";
                                            string LIMITDISBURSED = "";
                                            string ACC_NO = "";

                                            DataTable ABStatusDetailDT = conn.GetDataTable();
                                            conn.CloseConnection();

                                            DataRow[] ABStatusDetailROW = ABStatusDetailDT.Select("ID_AB_PARENT = '" + ID_PARENT + "'");

                                            for (int m = 0; m < ABStatusDetailROW.Length; m++)
                                            {
                                                string DESKRIPSI = ABStatusDetailROW[m]["DESCRIPT"].ToString();
                                                string LINENUM = ABStatusDetailROW[m]["LINENUM"].ToString();
                                                string QUERY_MARKING = ABStatusDetailROW[m]["QUERY_MARKING"].ToString();
                                                string ID_AB_STATUS_DETAIL = ABStatusDetailROW[m]["ID_AB_STATUS_DETAIL"].ToString();

                                                int post = 0;
                                                post = TOTAL_LINE + (int.Parse(LINENUM) - 1);
                                                string key = result[post];

                                                if (ID_AB_STATUS_DETAIL == "7" || ID_AB_STATUS_DETAIL == "4")
                                                {
                                                    PRODUCTID = key;
                                                }
                                                else if (ID_AB_STATUS_DETAIL == "8")
                                                {
                                                    ACC_SEQ = key;
                                                }

                                                infoTambahan = infoTambahan + DESKRIPSI + " : " + key + " - ";
                                            }

                                            string KODE_ERROR = result[result.Length - 2];
                                            string MESSAGE_ERROR = result[result.Length - 1];

                                            conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                            conn.QueryString = "INSERT INTO AB_ERROR_UPLOAD VALUES('" + regno + "', '" + KODE_ERROR + "', '" + MESSAGE_ERROR + "', '" + file + "', '" + infoTambahan + "');";
                                            conn.ExecuteQuery();
                                            conn.CloseConnection();
                                            error = true;

                                            if (READ_FEEDBACK != "")
                                            {
                                                if (ID_PARENT == "3")
                                                {
                                                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                                    conn.QueryString = "SELECT APPTYPE, CU_CIF, CP_LIMIT, PROD_SEQ FROM CUSTOMER, APPLICATION, CUSTPRODUCT, RFPRODUCT WHERE CUSTOMER.CU_REF = APPLICATION.CU_REF AND APPLICATION.AP_REGNO = CUSTPRODUCT.AP_REGNO AND CUSTPRODUCT.AP_REGNO = '" + regno + "' AND CUSTPRODUCT.PRODUCTID = RFPRODUCT.PRODUCTID AND RFPRODUCT.SIBS_PRODCODE = '" + PRODUCTID + "'";
                                                    conn.ExecuteQuery();
                                                    APPTYPE = conn.GetFieldValue(0, 0);
                                                    CIF = conn.GetFieldValue(0, 1);
                                                    LIMITDISBURSED = conn.GetFieldValue(0, 2);
                                                    PROD_SEQ = conn.GetFieldValue(0, 3);

                                                    if (preNameFile == "CIF")
                                                    {
                                                        /*
                                                             @AP_REGNO
                                                            @PRODUCTID
                                                            @APPTYPE
                                                            @USERID
                                                            @PROD_SEQ
                                                            @PG_TRACK
                                                         */
                                                        conn.QueryString = "exec TRACKUPDATE '" + regno + "', '" + PRODUCTID + "','" + APPTYPE + "','System','" + PROD_SEQ + "','BP18.0'";
                                                        conn.ExecuteQuery();
                                                    }
                                                    conn.CloseConnection();
                                                }
                                            }

                                            break;
                                        }
                                        else if (status == "N" && preNameFile == "COL")
                                        {
                                            //berarti dibalikin nunggu CIF uda kebentuk
                                            string KODE_ERROR = result[result.Length - 2];
                                            string MESSAGE_ERROR = result[result.Length - 1];

                                            if (KODE_ERROR != "" && MESSAGE_ERROR != "")
                                            {
                                                conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                                conn.QueryString = "INSERT INTO AB_ERROR_UPLOAD VALUES('" + regno + "', '" + KODE_ERROR + "', '" + MESSAGE_ERROR + "', '" + file + "', '');";
                                                conn.ExecuteQuery();
                                                conn.CloseConnection();
                                                error = true;
                                            }
                                            break;
                                        }
                                        else if (status == "")
                                        {
                                            break;
                                        }
                                    }

                                    if (k == (rowCount - 1) && error == false)
                                    {
                                        TOTAL_LINE = TOTAL_LINE + (int.Parse(ABStatusRow_LOOPINGCOUNTER) * rowCount);
                                        i = TOTAL_LINE;

                                        //dibawah ini must dibenerin
                                        int tot = 0;
                                        if (ID_PARENT == "2" || ID_PARENT == "3")
                                        {
                                            tot = 20;
                                        }
                                        else if (ID_PARENT == "6" )
                                        {
                                            tot = 30;
                                        }

                                        if (tot != 0)
                                        {
                                            int selisih = tot - rowCount;
                                            selisih = int.Parse(ABStatusRow_LOOPINGCOUNTER) * selisih;
                                            TOTAL_LINE = TOTAL_LINE + selisih;
                                            if(preNameFile == "CIF")
                                            {
                                                TOTAL_LINE = TOTAL_LINE + 1;
                                            }
                                            if (preNameFile == "COL")
                                            {
                                                TOTAL_LINE = TOTAL_LINE + 4;
                                            }
                                            i = TOTAL_LINE;
                                            // yang 1 adalah line counting
                                        }
                                    }

                                    if (error == true)
                                    {
                                        break;
                                    }

                                }

                                if (error == true)
                                {
                                    break;
                                }
                            }

                            if (error == true)
                            {
                                break;
                            }
                            else
                            {
                                conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                                conn.QueryString = "DELETE AB_ERROR_UPLOAD WHERE AP_REGNO = '" + regno + "'";
                                conn.ExecuteQuery();
                                conn.ClearData();
                                break;
                            }
                        }

                        voidIsUploadSucceed(regno, preNameFile);
                    }
                    //Console.WriteLine(line);

                    string newfile = file.Replace(ConfigurationManager.AppSettings["alfabitPathDownload"], ConfigurationManager.AppSettings["alfabitPathLog"]);
                    File.Delete(newfile);
                    File.Move(file, newfile);
                    File.Delete(file);
                }
                catch(Exception e)
                {
                    string m = e.Message;

                    StreamWriter a = File.CreateText("C:\\TESTERROR\\Test.txt");
                    a.Write(e.Message);
                    a.Close();
                }
            }
        }

        public static void voidIsUploadSucceed(string regno, string preNameFile)
        {
            //SUKSES : BP17.0, GAGAL : BP18.0
            //pertama cek aplikasi pake collateral atau gak
            //kalau gak, cek langsung ditable error, kalau kosong, update ke SUCCESS
            //klo iya dan table kosong dan tipenya CIF : invoke service pembuat collateral
            bool isError = false;
            conn = new Connection(ConfigurationManager.AppSettings["conn"]);
            conn.QueryString = "SELECT * FROM AB_ERROR_UPLOAD WHERE AP_REGNO = '" + regno + "'";
            conn.ExecuteQuery();
            if (conn.GetRowCount() > 0)
            {
                isError = true;
            }
            conn.CloseConnection();

            /*
             @AP_REGNO
            @PRODUCTID
            @APPTYPE
            @USERID
            @PROD_SEQ
            @PG_TRACK
             */

            //17.01 sukses
            //17.02 gagal

            if (preNameFile == "COL")
            {
                if (!isError)
                {
                    //check apakah ada prk
                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                    conn.QueryString = "SELECT COUNT(AP_REGNO) FROM CUSTPRODUCT CP, RFPRODUCT RP WHERE CP.PRODUCTID = RP.PRODUCTID AND (RP.IS_PRK = '1') AND (convert(varchar(20),CP.CP_NOTES) != 'LOAN DONE' or CP.CP_NOTES is null) AND AP_REGNO = '" + regno + "'";
                    conn.ExecuteQuery();
                    int count = int.Parse(conn.GetFieldValue(0,0));
                    conn.CloseConnection();

                    if (count > 0)
                    {
                        conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                        conn.QueryString = "INSERT INTO AB_ERROR_UPLOAD VALUES('" + regno + "', 'PRK', 'MENUNGGU PROSES UPLOAD CIF PRK', '', 'MENUNGGU PROSES UPLOAD CIF PRK')";
                        conn.ExecuteQuery();

                        UploadToCore U = new UploadToCore();
                        U.CreateUploadFile(regno, "CIF", "PRK");
                        conn.CloseConnection();
                    }
                    else
                    {
                        //update track jadi sukses
                        conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                        conn.QueryString = "UPDATE APPTRACK SET AP_CURRTRACK = 'BP17.0' WHERE AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();
                        conn.CloseConnection();
                    }
                }
                else
                {
                    //update track jadi failed
                    conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                    conn.QueryString = "UPDATE APPTRACK SET AP_CURRTRACK = 'BP18.0' WHERE AP_REGNO = '" + regno + "'";
                    conn.ExecuteQuery();
                    conn.CloseConnection();
                }
            }
            else if(preNameFile == "CIF")
            {
                conn = new Connection(ConfigurationManager.AppSettings["conn"]);
                conn.QueryString = "SELECT * FROM LISTCOLLATERAL WHERE AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    if (!isError)
                    {
                        conn.QueryString = "SELECT COUNT(AP_REGNO) FROM CUSTPRODUCT CP, RFPRODUCT RP WHERE CP.PRODUCTID = RP.PRODUCTID AND RP.IS_PRK = '1' AND (convert(varchar(20),CP.CP_NOTES) = 'LOAN DONE' or CP.CP_NOTES is null) AND AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();
                        int count = int.Parse(conn.GetFieldValue(0, 0));

                        if (count > 0)
                        {
                            conn.QueryString = "INSERT INTO AB_ERROR_UPLOAD VALUES('" + regno + "', 'COLLATERAL', 'MENUNGGU PROSES UPLOAD COLLATERAL PRK', '', 'MENUNGGU PROSES UPLOAD COLLATERAL')";
                            conn.ExecuteQuery();
                            UploadToCore U = new UploadToCore();
                            U.CreateUploadFile(regno, "COL", "PRK");
                        }
                        else
                        {
                            conn.QueryString = "INSERT INTO AB_ERROR_UPLOAD VALUES('" + regno + "', 'COLLATERAL', 'MENUNGGU PROSES UPLOAD COLLATERAL', '', 'MENUNGGU PROSES UPLOAD COLLATERAL')";
                            conn.ExecuteQuery();
                            UploadToCore U = new UploadToCore();
                            U.CreateUploadFile(regno, "COL", "");
                        }
                    }
                    else
                    {
                        //update jadi fail yg gak sukses
                        conn.QueryString = "UPDATE APPTRACK SET AP_CURRTRACK = 'BP18.0' WHERE AP_REGNO = '" + regno + "' AND AP_CURRTRACK <> 'BP17.0'";
                        conn.ExecuteQuery();
                    }
                }
                //klo g ada collateral
                else
                {
                    //klo g error
                    if (!isError)
                    {
                        conn.QueryString = "UPDATE APPTRACK SET AP_CURRTRACK = 'BP17.0' WHERE AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();
                    }
                }
                conn.CloseConnection();
            }
        }

        public static void Callback(object state)
        {
            Program.timer.Change(Timeout.Infinite, Timeout.Infinite);

            conn = new Connection(ConfigurationManager.AppSettings["conn"]);
            conn.QueryString = "SELECT ID_PARENT, COUNTING, READ_FEEDBACK, ID_AB_FILE FROM AB_PARENT_LOOPER ORDER BY ID_PARENT ASC";
            conn.ExecuteQuery();
            DataTable dtParentLooper = conn.GetDataTable();

            conn.QueryString = "SELECT [DESC] as DESCRIPT, LINENUM, LOOPINGCOUNTER, ID_PARENT FROM AB_STATUS ORDER BY ID_AB_STATUS ASC";
            conn.ExecuteQuery();
            DataTable dtABStatus = conn.GetDataTable();

            conn.QueryString = "SELECT ID_AB_PARENT, [DESC] as DESCRIPT, LINENUM, QUERY_MARKING FROM AB_STATUS_DETAIL ORDER BY ID_AB_STATUS_DETAIL ASC";
            conn.ExecuteQuery();
            DataTable dtABStatusDetail = conn.GetDataTable();

            conn.QueryString = "SELECT ID_AB_FILE, AB_FILE_NAME FROM AB_FILE ";
            conn.ExecuteQuery();
            DataTable dtABFile = conn.GetDataTable();
            conn.CloseConnection();

            Scanning(ConfigurationManager.AppSettings["alfabitPathDownload"], dtABFile, dtParentLooper, dtABStatus, dtABStatusDetail);

            Program.timer.Change(0, long.Parse(ConfigurationManager.AppSettings["AlphabitScanningDownloadFolder"]));
        }
    }
}
