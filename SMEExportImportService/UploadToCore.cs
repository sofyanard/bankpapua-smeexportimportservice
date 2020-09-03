using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using DMS.DBConnection;
using System.IO;
using System.Data;
using System.Collections;

namespace SMEExportImportService
{
    class UploadToCore : IUploadToCore
    {
        protected Connection conn = new Connection(ConfigurationManager.AppSettings["conn"]);

        public string CreateUploadFile(string regno, string type, string prk)
        {
            string RRN = "1";

            /********************************************** RRN ****************************************************/
            conn.QueryString = "SELECT BRANCH_CODE FROM APPLICATION WHERE AP_REGNO = '" + regno + "'";
            conn.ExecuteQuery();
            string branchCODE = conn.GetFieldValue("BRANCH_CODE");

            conn.QueryString = "DELETE AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
            conn.ExecuteQuery();

            conn.QueryString = "SELECT RRN FROM AB_USED_RRN WHERE AP_REGNO = '" + regno + "'";
            conn.ExecuteQuery();

            DataTable dtRRN = conn.GetDataTable().Copy();
            int rrn = dtRRN.Rows.Count;
            if (dtRRN.Rows.Count == 0)
            {
                string nol = "";
                if ((nol + regno).Length < 20)
                {
                    int selisih = 20 - regno.Length;
                    for (int a = 0; a <= selisih; a++)
                    {
                        nol = nol + "0";
                        if ((nol + regno).Length == 20)
                        {
                            break;
                        }
                    }
                }

                conn.QueryString = "INSERT INTO AB_USED_RRN VALUES ('" + regno + "', '" + nol + regno + "', '" + rrn + "')";
                conn.ExecuteQuery();
            }
            else 
            {
                string nol = "";
                if ((nol + regno).Length < 20)
                {
                    int selisih = 20 - regno.Length;
                    for (int a = 0; a <= selisih; a++)
                    {
                        nol = nol + "0";
                        if ((nol + regno).Length == 20)
                        {
                            break;
                        }
                    }
                }

                for (int i = 0; i < rrn.ToString().Length; i++)
                {
                    nol = nol.Remove(nol.Length - 1, 1);
                }

                nol = nol + rrn.ToString();

                conn.QueryString = "INSERT INTO AB_USED_RRN VALUES ('" + regno + "', '" + nol + regno + "', '" + rrn + "')";
                conn.ExecuteQuery();
            }

            conn.QueryString = "SELECT RRN FROM AB_USED_RRN WHERE AP_REGNO = '" + regno + "' AND RRN_SEQ = '" + rrn + "'";
            conn.ExecuteQuery();
            dtRRN = conn.GetDataTable().Copy();
            RRN = dtRRN.Rows[0][0].ToString().Trim();
            /*******************************************************************************************************/

            conn.QueryString = "SELECT AB_FILE_NAME, ID_AB_FILE FROM AB_FILE ORDER BY ID_AB_FILE ASC";
            conn.ExecuteQuery();

            DataTable dt1 = conn.GetDataTable().Copy();
            DataRow[] dataABFile = dt1.Select("AB_FILE_NAME = '" + type + "'");

            if (dataABFile.Length > 0)
            {
                Dictionary<string, ArrayList> fileDictionary = new Dictionary<string, ArrayList>();
                for (int i = 0; i < dataABFile.Length; i++)
                {
                    int postFileName = 0;
                    string FILENAME = dataABFile[i][0].ToString().Trim();
                    string idFILE = dataABFile[i][1].ToString().Trim();

                    if (prk == "")
                    {
                        conn.QueryString = "SELECT ID_PARENT, QUERY, [SELECT], COUNTING FROM AB_PARENT_LOOPER WHERE ID_AB_FILE = '" + idFILE + "' AND COUNTING <> '' ORDER BY ID_PARENT ASC";
                        conn.ExecuteQuery();
                    }
                    else 
                    {
                        conn.QueryString = "SELECT ID_PARENT, QUERY, [SELECT], COUNTING FROM AB_PARENT_LOOPER_PRK WHERE ID_AB_FILE = '" + idFILE + "' AND COUNTING <> '' ORDER BY ID_PARENT ASC";
                        conn.ExecuteQuery();
                    }

                    DataTable dt2 = conn.GetDataTable().Copy();

                    if (dt2.Rows.Count > 0)
                    {
                        for (int j = 0; j < dt2.Rows.Count; j++)
                        {
                            string ID_PARENT = dt2.Rows[j][0].ToString().Trim();
                            string QUERY = dt2.Rows[j][1].ToString().Trim();
                            string SELECT = dt2.Rows[j][2].ToString().Trim();
                            string COUNTING = dt2.Rows[j][3].ToString().Trim();

                            int counter = 0;
                            try
                            {
                                COUNTING = COUNTING.Replace("#AP_REGNO", "'" + regno + "'");
                                conn.QueryString = COUNTING;
                                conn.ExecuteQuery();
                                counter = int.Parse(conn.GetFieldValue(0, 0));
                            }
                            catch
                            {

                            }

                            //dapatkan selector
                            string[] selector = SELECT.Split('#');

                            QUERY = QUERY.Replace("#AP_REGNO", "'" + regno + "'");

                            conn.QueryString = QUERY;
                            conn.ExecuteQuery();

                            ArrayList parentLooper = new ArrayList();
                            parentLooper.Clear();
                            //looping sesuai jumlah rownya
                            for (int z = 0; z < conn.GetRowCount(); z++)
                            {
                                Dictionary<string, string> dictionary = new Dictionary<string, string>();
                                for (int k = 0; k < selector.Length; k++)
                                {
                                    dictionary.Add(selector[k], conn.GetFieldValue(z, selector[k]));
                                }
                                parentLooper.Add(dictionary);
                            }

                            postFileName = postFileName + 1;
                            fileDictionary.Add(ID_PARENT, parentLooper);
                        }
                    }
                }

                if (prk == "")
                {
                    splitToTwenty(fileDictionary, branchCODE, regno, type);
                }
                else
                {
                    splitToTwentyPRK(fileDictionary, branchCODE, regno, type);
                }
            }

            return "Sukses !";
        }

        public void splitToTwenty(Dictionary<string, ArrayList> fileDictionary, string branchCODE, string regno, string type)
        {
            string urutanRRN_str = "";
            int urutanFILE = 1;
            int urutanRRN = 1;
            string RRN = "";

            conn.QueryString = "SELECT ID_AB_FILE FROM AB_FILE WHERE AB_FILE_NAME = '" + type + "'";
            conn.ExecuteQuery();

            string types = conn.GetFieldValue("ID_AB_FILE");

            conn.QueryString = "SELECT ID_PARENT, ID_AB_FILE FROM AB_PARENT_LOOPER WHERE COUNTING <> '' AND ID_AB_FILE = ' " + types + "' ORDER BY ID_PARENT ASC";
            conn.ExecuteQuery();
            DataTable AB_PARENT_LOOPER = conn.GetDataTable();

            string ID_AB_FILE_PREV = "";
            StreamWriter yourStream = null;
            bool breakloop = false;

            while (true)
            {
                for (int i = 0; i < AB_PARENT_LOOPER.Rows.Count; i++)
                {
                    //bikin file baru
                    string ID_PARENT = AB_PARENT_LOOPER.Rows[i]["ID_PARENT"].ToString().Trim();
                    string ID_AB_FILE = AB_PARENT_LOOPER.Rows[i]["ID_AB_FILE"].ToString().Trim();

                    conn.QueryString = "SELECT AB_FILE_NAME FROM AB_FILE WHERE ID_AB_FILE = '" + ID_AB_FILE + "'";
                    conn.ExecuteQuery();

                    string FILENAME = conn.GetFieldValue("AB_FILE_NAME");

                    conn.QueryString = "SELECT QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH FROM AB_CHILDREN WHERE ID_PARENT = " + ID_PARENT + " ORDER BY ID_CHILDREN ASC";
                    conn.ExecuteQuery();
                    DataTable dtAB_CHILD = conn.GetDataTable().Copy();

                    //klo ab file uda beda baru bikin stream baru;
                    if (ID_AB_FILE != ID_AB_FILE_PREV)
                    {
                        try
                        {
                            if (ID_AB_FILE_PREV != "")
                            {
                                
                            }
                            yourStream.Close();
                        }
                        catch
                        {

                        }

                        conn.QueryString = "SELECT URUTAN_RRN_5_DIGIT, [FILE] as F FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();

                        if (conn.GetRowCount() > 0)
                        {
                            conn.QueryString = "SELECT MAX(URUTAN_RRN) as URUTAN FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();

                            try
                            {
                                urutanRRN = int.Parse(conn.GetFieldValue("URUTAN"));

                            }
                            catch
                            {
                                urutanRRN = -1;
                            }

                            urutanRRN = urutanRRN + 1;

                            string urutanString = "";
                            urutanString = urutanRRN.ToString();

                            if (urutanString.Length < 5)
                            {
                                urutanString = urutanString.PadLeft(5, '0');
                            }

                            conn.QueryString = "SELECT MAX(URUTAN_FILE) FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "'";
                            conn.ExecuteQuery();

                            try
                            {
                                urutanFILE = int.Parse(conn.GetFieldValue(0, 0));
                                urutanFILE += 1;

                                if (urutanFILE == 9999)
                                {
                                    urutanFILE = 1;
                                }
                            }
                            catch
                            {
                                urutanFILE = 1;
                            }

                            conn.QueryString = "INSERT INTO AB_URUTAN VALUES('" + branchCODE + "','" + regno + "', '" + FILENAME + branchCODE + urutanFILE.ToString() + ".txt" + "'," + urutanFILE + "," + urutanRRN + ",'" + urutanString + "')";
                            conn.ExecuteNonQuery();

                            conn.QueryString = "SELECT URUTAN_RRN_5_DIGIT, [FILE] as F FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();
                        }
                        else
                        {
                            conn.QueryString = "SELECT MAX(URUTAN_FILE) FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "'";
                            conn.ExecuteQuery();

                            try
                            {
                                urutanFILE = int.Parse(conn.GetFieldValue(0, 0));
                                urutanFILE += 1;

                                if (urutanFILE == 9999)
                                {
                                    urutanFILE = 1;
                                }
                            }
                            catch
                            {
                                urutanFILE = 1;
                            }

                            conn.QueryString = "INSERT INTO AB_URUTAN VALUES('" + branchCODE + "','" + regno + "', '" + FILENAME + branchCODE + urutanFILE.ToString() + ".txt" + "'," + urutanFILE + "," + 1 + ",'00001')";
                            conn.ExecuteNonQuery();

                            conn.QueryString = "SELECT URUTAN_RRN_5_DIGIT, [FILE] as F FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();
                        }

                        urutanRRN_str = conn.GetFieldValue("URUTAN_RRN_5_DIGIT");
                        yourStream = File.CreateText(ConfigurationManager.AppSettings["alfabitPathUpload"] + "\\" + conn.GetFieldValue("F")); // creating file

                        //getRRN
                        conn.QueryString = "SELECT RRN FROM AB_USED_RRN WHERE AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();

                        RRN = conn.GetFieldValue("RRN");
                        FillHeader(yourStream, regno, ID_AB_FILE);
                    }

                    int lastrow = 0;
                    ArrayList isi = fileDictionary[ID_PARENT];

                    int max = 0;
                    //CIF
                    if (ID_AB_FILE == "1")
                    {
                        max = 20;
                    }
                    //COLL
                    else if(ID_AB_FILE == "2")
                    {
                        max = 30;
                    }

                    bool COUNTER_IS_WRITTEN = false;
                    for (int j = 0; j < max; j++)
                    {
                        try
                        {
                            /*
                            QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH
                             */

                            Dictionary<string, string> dict = (Dictionary<string, string>)isi[j];
                            //QUERYC = QUERYC.Replace(KeysReplaced, "'" + dict[Keys] + "'");
                            for (int m = 0; m < dtAB_CHILD.Rows.Count; m++)
                            {
                                string QUERY_CHILD = dtAB_CHILD.Rows[m]["QUERY_CHILD"].ToString().Trim();
                                string ID_CHILDREN = dtAB_CHILD.Rows[m]["ID_CHILDREN"].ToString().Trim();
                                string DATA_DEC = dtAB_CHILD.Rows[m]["DATA_DEC"].ToString().Trim();
                                string DATA_JENIS = dtAB_CHILD.Rows[m]["DATA_JENIS"].ToString().Trim();
                                string DATA_LENGTH = dtAB_CHILD.Rows[m]["DATA_LENGTH"].ToString().Trim();

                                for (int l = 0; l < dict.Count; l++)
                                {
                                    string Keys = dict.Keys.ElementAt(l);
                                    string KeysReplaced = "#" + Keys;
                                    //dtAB_CHILD.Rows

                                    QUERY_CHILD = QUERY_CHILD.Replace(KeysReplaced, "'" + dict[Keys] + "'"); 
                                }

                                if (QUERY_CHILD != "")
                                 {
                                    conn.QueryString = QUERY_CHILD;
                                    conn.ExecuteQuery();

                                    string content = conn.GetFieldValue("CONTENT");

                                    if(DATA_JENIS == "N")
                                    {
                                        if (DATA_DEC != "")
                                        {
                                            int DEC = int.Parse(DATA_DEC);
                                            int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                                            string[] contents = content.Split(new char[] {','});

                                            if (contents.Length == 1)
                                            {
                                                if (DEC > 0)
                                                {
                                                    content = contents[0].PadLeft(SELISIH, '0');
                                                    content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                                }
                                                else
                                                {
                                                    content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                                }
                                            }
                                            else 
                                            {
                                                string content1 = contents[0];
                                                string content2 = contents[1];

                                                content1 = content1.Replace(".", "");
                                                content1 = content1.PadLeft(SELISIH, '0');
                                                content2 = content2.PadRight(DEC, '0');

                                                content = content1 + content2;
                                            }
                                        }
                                        else
                                        {
                                            content = content.Replace(".", "");
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }
                                    else if(DATA_JENIS == "D")
                                    {
                                        try
                                        {
                                            DateTime dt = DateTime.Parse(content);
                                            string date = String.Format("{0:dd-MM-yyyy}", dt);
                                            content = date.Replace("-", "");
                                        }
                                        catch
                                        {
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }

                                    yourStream.Write(content + Environment.NewLine);
                                }
                                else if (QUERY_CHILD == "" && COUNTER_IS_WRITTEN == false)
                                {
                                    COUNTER_IS_WRITTEN = true;
                                    conn.QueryString = "SELECT COUNTING FROM AB_PARENT_LOOPER WHERE ID_PARENT = '" + ID_PARENT + "'";
                                    conn.ExecuteQuery();

                                    string counting = conn.GetFieldValue(0, 0);

                                    conn.QueryString = counting.Replace("#AP_REGNO", "'" + regno + "'");
                                    conn.ExecuteQuery();

                                    string content = conn.GetFieldValue(0, 0);

                                    if (DATA_JENIS == "N")
                                    {
                                        if (DATA_DEC != "")
                                        {
                                            int DEC = int.Parse(DATA_DEC);
                                            int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                                            string[] contents = content.Split(new char[] { ',' });

                                            if (contents.Length == 1)
                                            {
                                                if (DEC > 0)
                                                {
                                                    content = contents[0].PadLeft(SELISIH, '0');
                                                    content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                                }
                                                else
                                                {
                                                    content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                                }
                                            }
                                            else
                                            {
                                                string content1 = contents[0];
                                                string content2 = contents[1];

                                                content1 = content1.Replace(".", "");
                                                content1 = content1.PadLeft(SELISIH, '0');
                                                content2 = content2.PadRight(DEC, '0');

                                                content = content1 + content2;
                                            }
                                        }
                                        else
                                        {
                                            content = content.Replace(".", "");
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }
                                    else if (DATA_JENIS == "D")
                                    {
                                        try
                                        {
                                            DateTime dt = DateTime.Parse(content);
                                            string date = String.Format("{0:dd-MM-yyyy}", dt);
                                            content = date.Replace("-", "");
                                        }
                                        catch
                                        {
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }

                                    yourStream.Write(content + Environment.NewLine);
                                }
                            }
                        }
                        catch
                        {
                            try
                            {
                                conn.QueryString = "SELECT DISTINCT(LOOPINGCOUNTER) as CONTENT FROM AB_STATUS WHERE ID_PARENT = '" + ID_PARENT + "'";
                                conn.ExecuteQuery();

                                int rows = int.Parse(conn.GetFieldValue("CONTENT"));

                                for (int k = 0; k < rows; k++)
                                {
                                    yourStream.Write("" + Environment.NewLine);
                                }
                            }
                            catch
                            {

                            }
                        }

                        if (j == (max-1))
                        {
                            //testing purpose
                            lastrow = max;
                        }
                    }
                    //replace isi, kan uda ditulis
                    fileDictionary[ID_PARENT].Clear();
                    fileDictionary[ID_PARENT].Add(isi);

                    ID_AB_FILE_PREV = ID_AB_FILE;
                    if (lastrow == max)
                    {
                        //jika 20 cek ID_PARENT yang ID_AB_FILE nya sama
                        //klo dia yang paling maks, create text baru
                        //klo bukan g perlu bikin teks baru

                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE ID_AB_FILE = '" + ID_AB_FILE + "'";
                        conn.ExecuteQuery();

                        int ID_PARENT_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE COUNTING <> '' AND ID_AB_FILE = '" + types + "'";
                        conn.ExecuteQuery();

                        int ID_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        if (ID_PARENT.ToString() == ID_MAX.ToString())
                        {
                            FillFooter(yourStream, regno, ID_AB_FILE);
                            yourStream.Close();
                            breakloop = true;
                        }
                        else if (ID_PARENT_MAX.ToString() == ID_PARENT.ToString())
                        {
                            ID_AB_FILE_PREV = "";
                        }
                    }
                    else
                    {
                        //jika bukan 20 cek ID_PARENT yang ID_AB_FILE nya sama
                        //klo dia yang paling maks break
                        //klo bukan CONTINUE
                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE ID_AB_FILE = '" + ID_AB_FILE + "'";
                        conn.ExecuteQuery();

                        int ID_PARENT_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE COUNTING <> ''";
                        conn.ExecuteQuery();

                        int ID_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        if(ID_PARENT.ToString() == ID_MAX.ToString())
                        {
                            yourStream.Close();
                            FillFooter(yourStream, regno, ID_AB_FILE);
                            breakloop = true;
                        }
                        else if (ID_PARENT_MAX.ToString() == ID_PARENT.ToString())
                        {
                            continue;
                        }
                    }
                }

                if(breakloop == true)
                {
                    break;
                }
            }
        }

        public void splitToTwentyPRK(Dictionary<string, ArrayList> fileDictionary, string branchCODE, string regno, string type)
        {
            string urutanRRN_str = "";
            int urutanFILE = 1;
            int urutanRRN = 1;
            string RRN = "";

            conn.QueryString = "SELECT ID_AB_FILE FROM AB_FILE WHERE AB_FILE_NAME = '" + type + "'";
            conn.ExecuteQuery();

            string types = conn.GetFieldValue("ID_AB_FILE");

            conn.QueryString = "SELECT ID_PARENT, ID_AB_FILE FROM AB_PARENT_LOOPER_PRK WHERE COUNTING <> '' AND ID_AB_FILE = ' " + types + "' ORDER BY ID_PARENT ASC";
            conn.ExecuteQuery();
            DataTable AB_PARENT_LOOPER = conn.GetDataTable();

            string ID_AB_FILE_PREV = "";
            StreamWriter yourStream = null;
            bool breakloop = false;

            while (true)
            {
                for (int i = 0; i < AB_PARENT_LOOPER.Rows.Count; i++)
                {
                    //bikin file baru
                    string ID_PARENT = AB_PARENT_LOOPER.Rows[i]["ID_PARENT"].ToString().Trim();
                    string ID_AB_FILE = AB_PARENT_LOOPER.Rows[i]["ID_AB_FILE"].ToString().Trim();

                    conn.QueryString = "SELECT AB_FILE_NAME FROM AB_FILE WHERE ID_AB_FILE = '" + ID_AB_FILE + "'";
                    conn.ExecuteQuery();

                    string FILENAME = conn.GetFieldValue("AB_FILE_NAME");

                    conn.QueryString = "SELECT QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH FROM AB_CHILDREN_PRK WHERE ID_PARENT = " + ID_PARENT + " ORDER BY ID_CHILDREN ASC";
                    conn.ExecuteQuery();
                    DataTable dtAB_CHILD = conn.GetDataTable().Copy();

                    //klo ab file uda beda baru bikin stream baru;
                    if (ID_AB_FILE != ID_AB_FILE_PREV)
                    {
                        try
                        {
                            if (ID_AB_FILE_PREV != "")
                            {

                            }
                            yourStream.Close();
                        }
                        catch
                        {

                        }

                        conn.QueryString = "SELECT URUTAN_RRN_5_DIGIT, [FILE] as F FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();

                        if (conn.GetRowCount() > 0)
                        {
                            conn.QueryString = "SELECT MAX(URUTAN_RRN) as URUTAN FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();

                            try
                            {
                                urutanRRN = int.Parse(conn.GetFieldValue("URUTAN"));

                            }
                            catch
                            {
                                urutanRRN = -1;
                            }

                            urutanRRN = urutanRRN + 1;

                            string urutanString = "";
                            urutanString = urutanRRN.ToString();

                            if (urutanString.Length < 5)
                            {
                                urutanString = urutanString.PadLeft(5, '0');
                            }

                            conn.QueryString = "SELECT MAX(URUTAN_FILE) FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "'";
                            conn.ExecuteQuery();

                            try
                            {
                                urutanFILE = int.Parse(conn.GetFieldValue(0, 0));
                                urutanFILE += 1;

                                if (urutanFILE == 9999)
                                {
                                    urutanFILE = 1;
                                }
                            }
                            catch
                            {
                                urutanFILE = 1;
                            }

                            conn.QueryString = "INSERT INTO AB_URUTAN VALUES('" + branchCODE + "','" + regno + "', '" + FILENAME + branchCODE + urutanFILE.ToString() + ".txt" + "'," + urutanFILE + "," + urutanRRN + ",'" + urutanString + "')";
                            conn.ExecuteNonQuery();

                            conn.QueryString = "SELECT URUTAN_RRN_5_DIGIT, [FILE] as F FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();
                        }
                        else
                        {
                            conn.QueryString = "SELECT MAX(URUTAN_FILE) FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "'";
                            conn.ExecuteQuery();

                            try
                            {
                                urutanFILE = int.Parse(conn.GetFieldValue(0, 0));
                                urutanFILE += 1;

                                if (urutanFILE == 9999)
                                {
                                    urutanFILE = 1;
                                }
                            }
                            catch
                            {
                                urutanFILE = 1;
                            }

                            conn.QueryString = "INSERT INTO AB_URUTAN VALUES('" + branchCODE + "','" + regno + "', '" + FILENAME + branchCODE + urutanFILE.ToString() + ".txt" + "'," + urutanFILE + "," + 1 + ",'00001')";
                            conn.ExecuteNonQuery();

                            conn.QueryString = "SELECT URUTAN_RRN_5_DIGIT, [FILE] as F FROM AB_URUTAN WHERE BRANCH_CODE = '" + branchCODE + "' AND AP_REGNO = '" + regno + "'";
                            conn.ExecuteQuery();
                        }

                        urutanRRN_str = conn.GetFieldValue("URUTAN_RRN_5_DIGIT");
                        yourStream = File.CreateText(ConfigurationManager.AppSettings["alfabitPathUpload"] + "\\" + conn.GetFieldValue("F")); // creating file

                        //getRRN
                        conn.QueryString = "SELECT RRN FROM AB_USED_RRN WHERE AP_REGNO = '" + regno + "'";
                        conn.ExecuteQuery();

                        RRN = conn.GetFieldValue("RRN");
                        FillHeaderPRK(yourStream, regno, ID_AB_FILE);
                    }

                    int lastrow = 0;
                    ArrayList isi = fileDictionary[ID_PARENT];

                    int max = 0;
                    //CIF
                    if (ID_AB_FILE == "1")
                    {
                        max = 20;
                    }
                    //COLL
                    else if (ID_AB_FILE == "2")
                    {
                        max = 30;
                    }

                    bool COUNTER_IS_WRITTEN = false;
                    for (int j = 0; j < max; j++)
                    {
                        try
                        {
                            /*
                            QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH
                             */

                            Dictionary<string, string> dict = (Dictionary<string, string>)isi[j];
                            //QUERYC = QUERYC.Replace(KeysReplaced, "'" + dict[Keys] + "'");
                            for (int m = 0; m < dtAB_CHILD.Rows.Count; m++)
                            {
                                string QUERY_CHILD = dtAB_CHILD.Rows[m]["QUERY_CHILD"].ToString().Trim();
                                string ID_CHILDREN = dtAB_CHILD.Rows[m]["ID_CHILDREN"].ToString().Trim();
                                string DATA_DEC = dtAB_CHILD.Rows[m]["DATA_DEC"].ToString().Trim();
                                string DATA_JENIS = dtAB_CHILD.Rows[m]["DATA_JENIS"].ToString().Trim();
                                string DATA_LENGTH = dtAB_CHILD.Rows[m]["DATA_LENGTH"].ToString().Trim();

                                for (int l = 0; l < dict.Count; l++)
                                {
                                    string Keys = dict.Keys.ElementAt(l);
                                    string KeysReplaced = "#" + Keys;
                                    //dtAB_CHILD.Rows

                                    QUERY_CHILD = QUERY_CHILD.Replace(KeysReplaced, "'" + dict[Keys] + "'");
                                }

                                if (QUERY_CHILD != "")
                                {
                                    conn.QueryString = QUERY_CHILD;
                                    conn.ExecuteQuery();

                                    string content = conn.GetFieldValue("CONTENT");

                                    if (DATA_JENIS == "N")
                                    {
                                        if (DATA_DEC != "")
                                        {
                                            int DEC = int.Parse(DATA_DEC);
                                            int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                                            string[] contents = content.Split(new char[] { ',' });

                                            if (contents.Length == 1)
                                            {
                                                if (DEC > 0)
                                                {
                                                    content = contents[0].PadLeft(SELISIH, '0');
                                                    content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                                }
                                                else
                                                {
                                                    content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                                }
                                            }
                                            else
                                            {
                                                string content1 = contents[0];
                                                string content2 = contents[1];

                                                content1 = content1.Replace(".", "");
                                                content1 = content1.PadLeft(SELISIH, '0');
                                                content2 = content2.PadRight(DEC, '0');

                                                content = content1 + content2;
                                            }
                                        }
                                        else
                                        {
                                            content = content.Replace(".", "");
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }
                                    else if (DATA_JENIS == "D")
                                    {
                                        try
                                        {
                                            DateTime dt = DateTime.Parse(content);
                                            string date = String.Format("{0:dd-MM-yyyy}", dt);
                                            content = date.Replace("-", "");
                                        }
                                        catch
                                        {
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }

                                    yourStream.Write(content + Environment.NewLine);
                                }
                                else if (QUERY_CHILD == "" && COUNTER_IS_WRITTEN == false)
                                {
                                    COUNTER_IS_WRITTEN = true;
                                    conn.QueryString = "SELECT COUNTING FROM AB_PARENT_LOOPER_PRK WHERE ID_PARENT = '" + ID_PARENT + "'";
                                    conn.ExecuteQuery();

                                    string counting = conn.GetFieldValue(0, 0);

                                    conn.QueryString = counting.Replace("#AP_REGNO", "'" + regno + "'");
                                    conn.ExecuteQuery();

                                    string content = conn.GetFieldValue(0, 0);

                                    if (DATA_JENIS == "N")
                                    {
                                        if (DATA_DEC != "")
                                        {
                                            int DEC = int.Parse(DATA_DEC);
                                            int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                                            string[] contents = content.Split(new char[] { ',' });

                                            if (contents.Length == 1)
                                            {
                                                if (DEC > 0)
                                                {
                                                    content = contents[0].PadLeft(SELISIH, '0');
                                                    content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                                }
                                                else
                                                {
                                                    content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                                }
                                            }
                                            else
                                            {
                                                string content1 = contents[0];
                                                string content2 = contents[1];

                                                content1 = content1.Replace(".", "");
                                                content1 = content1.PadLeft(SELISIH, '0');
                                                content2 = content2.PadRight(DEC, '0');

                                                content = content1 + content2;
                                            }
                                        }
                                        else
                                        {
                                            content = content.Replace(".", "");
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }
                                    else if (DATA_JENIS == "D")
                                    {
                                        try
                                        {
                                            DateTime dt = DateTime.Parse(content);
                                            string date = String.Format("{0:dd-MM-yyyy}", dt);
                                            content = date.Replace("-", "");
                                        }
                                        catch
                                        {
                                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                                        }
                                    }

                                    yourStream.Write(content + Environment.NewLine);
                                }
                            }
                        }
                        catch
                        {
                            try
                            {
                                conn.QueryString = "SELECT DISTINCT(LOOPINGCOUNTER) as CONTENT FROM AB_STATUS WHERE ID_PARENT = '" + ID_PARENT + "'";
                                conn.ExecuteQuery();

                                int rows = int.Parse(conn.GetFieldValue("CONTENT"));

                                for (int k = 0; k < rows; k++)
                                {
                                    yourStream.Write("" + Environment.NewLine);
                                }
                            }
                            catch
                            {

                            }
                        }

                        if (j == (max - 1))
                        {
                            //testing purpose
                            lastrow = max;
                        }
                    }
                    //replace isi, kan uda ditulis
                    fileDictionary[ID_PARENT].Clear();
                    fileDictionary[ID_PARENT].Add(isi);

                    ID_AB_FILE_PREV = ID_AB_FILE;
                    if (lastrow == max)
                    {
                        //jika 20 cek ID_PARENT yang ID_AB_FILE nya sama
                        //klo dia yang paling maks, create text baru
                        //klo bukan g perlu bikin teks baru

                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE ID_AB_FILE = '" + ID_AB_FILE + "'";
                        conn.ExecuteQuery();

                        int ID_PARENT_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE COUNTING <> '' AND ID_AB_FILE = '" + types + "'";
                        conn.ExecuteQuery();

                        int ID_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        if (ID_PARENT.ToString() == ID_MAX.ToString())
                        {
                            FillFooterPRK(yourStream, regno, ID_AB_FILE);
                            yourStream.Close();
                            breakloop = true;
                        }
                        else if (ID_PARENT_MAX.ToString() == ID_PARENT.ToString())
                        {
                            ID_AB_FILE_PREV = "";
                        }
                    }
                    else
                    {
                        //jika bukan 20 cek ID_PARENT yang ID_AB_FILE nya sama
                        //klo dia yang paling maks break
                        //klo bukan CONTINUE
                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE ID_AB_FILE = '" + ID_AB_FILE + "'";
                        conn.ExecuteQuery();

                        int ID_PARENT_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE COUNTING <> ''";
                        conn.ExecuteQuery();

                        int ID_MAX = int.Parse(conn.GetFieldValue("ID_PARENT"));

                        if (ID_PARENT.ToString() == ID_MAX.ToString())
                        {
                            yourStream.Close();
                            FillFooterPRK(yourStream, regno, ID_AB_FILE);
                            breakloop = true;
                        }
                        else if (ID_PARENT_MAX.ToString() == ID_PARENT.ToString())
                        {
                            continue;
                        }
                    }
                }

                if (breakloop == true)
                {
                    break;
                }
            }
        }

        public void FillFooter(StreamWriter stream, string regno, string ID_AB_FILE)
        {
            conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE COUNTING = '' AND ID_AB_FILE = '" + ID_AB_FILE + "'";
            conn.ExecuteQuery();

            string ID_PARENT_MAX = conn.GetFieldValue(0, 0);

            conn.QueryString = "SELECT MIN(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE COUNTING = '' AND ID_AB_FILE = '" + ID_AB_FILE + "'";
            conn.ExecuteQuery();

            string ID_PARENT_MIN = conn.GetFieldValue(0, 0);

            try
            {
                if (ID_PARENT_MAX != ID_PARENT_MIN)
                {
                    conn.QueryString = "SELECT QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH FROM AB_CHILDREN WHERE ID_PARENT = " + ID_PARENT_MAX + " ORDER BY ID_CHILDREN ASC";
                    conn.ExecuteQuery();
                    DataTable dt3 = conn.GetDataTable().Copy();

                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        string QUERYC = dt3.Rows[i]["QUERY_CHILD"].ToString().Trim();
                        string ID_CHILDREN = dt3.Rows[i]["ID_CHILDREN"].ToString().Trim();
                        string DATA_DEC = dt3.Rows[i]["DATA_DEC"].ToString().Trim();
                        string DATA_JENIS = dt3.Rows[i]["DATA_JENIS"].ToString().Trim();
                        string DATA_LENGTH = dt3.Rows[i]["DATA_LENGTH"].ToString().Trim();

                        QUERYC = QUERYC.Replace("#AP_REGNO", "'" + regno + "'");
                        conn.QueryString = QUERYC;
                        conn.ExecuteQuery();

                        string content = conn.GetFieldValue("CONTENT");

                        if (DATA_JENIS == "N")
                        {
                            if (DATA_DEC != "")
                            {
                                int DEC = int.Parse(DATA_DEC);
                                int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                                string[] contents = content.Split(new char[] { ',' });

                                if (contents.Length == 1)
                                {
                                    if (DEC > 0)
                                    {
                                        content = contents[0].PadLeft(SELISIH, '0');
                                        content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                    }
                                    else
                                    {
                                        content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                    }
                                }
                                else
                                {
                                    string content1 = contents[0];
                                    string content2 = contents[1];

                                    content1 = content1.Replace(".", "");
                                    content1 = content1.PadLeft(SELISIH, '0');
                                    content2 = content2.PadRight(DEC, '0');

                                    content = content1 + content2;
                                }
                            }
                            else
                            {
                                content = content.Replace(".", "");
                                content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                            }
                        }
                        else if (DATA_JENIS == "D")
                        {
                            try
                            {
                                DateTime dt = DateTime.Parse(content);
                                string date = String.Format("{0:dd-MM-yyyy}", dt);
                                content = date.Replace("-", "");
                            }
                            catch
                            {
                                content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                            }
                        }

                        stream.Write(content + Environment.NewLine);
                    }
                }
            }
            catch
            {

            }
        }

        public void FillHeader(StreamWriter stream, string regno, string ID_AB_FILE)
        {
            conn.QueryString = "SELECT MIN(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER WHERE COUNTING = '' AND ID_AB_FILE = '" + ID_AB_FILE + "'";
            conn.ExecuteQuery();

            try
            {
                string ID_PARENT = conn.GetFieldValue(0, 0);

                conn.QueryString = "SELECT QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH FROM AB_CHILDREN WHERE ID_PARENT = " + ID_PARENT + " ORDER BY ID_CHILDREN ASC";
                conn.ExecuteQuery();
                DataTable dt3 = conn.GetDataTable().Copy();

                for (int i = 0; i < dt3.Rows.Count; i++)
                {
                    string QUERYC = dt3.Rows[i]["QUERY_CHILD"].ToString().Trim();
                    string ID_CHILDREN = dt3.Rows[i]["ID_CHILDREN"].ToString().Trim();
                    string DATA_DEC = dt3.Rows[i]["DATA_DEC"].ToString().Trim();
                    string DATA_JENIS = dt3.Rows[i]["DATA_JENIS"].ToString().Trim();
                    string DATA_LENGTH = dt3.Rows[i]["DATA_LENGTH"].ToString().Trim();

                    QUERYC = QUERYC.Replace("#AP_REGNO", "'" + regno + "'");
                    conn.QueryString = QUERYC;
                    conn.ExecuteQuery();

                    string content = conn.GetFieldValue("CONTENT");

                    if (DATA_JENIS == "N")
                    {
                        if (DATA_DEC != "")
                        {
                            int DEC = int.Parse(DATA_DEC);
                            int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                            string[] contents = content.Split(new char[] { ',' });

                            if (contents.Length == 1)
                            {
                                if (DEC > 0)
                                {
                                    content = contents[0].PadLeft(SELISIH, '0');
                                    content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                }
                                else
                                {
                                    content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                }
                            }
                            else
                            {
                                string content1 = contents[0];
                                string content2 = contents[1];

                                content1 = content1.Replace(".", "");
                                content1 = content1.PadLeft(SELISIH, '0');
                                content2 = content2.PadRight(DEC, '0');

                                content = content1 + content2;
                            }
                        }
                        else
                        {
                            content = content.Replace(".", "");
                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                        }
                    }
                    else if (DATA_JENIS == "D")
                    {
                        try
                        {
                            DateTime dt = DateTime.Parse(content);
                            string date = String.Format("{0:dd-MM-yyyy}", dt);
                            content = date.Replace("-", "");
                        }
                        catch
                        {
                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                        }
                    }

                    stream.Write(content + Environment.NewLine);
                }
            }
            catch
            {

            }
        }

        public void FillFooterPRK(StreamWriter stream, string regno, string ID_AB_FILE)
        {
            conn.QueryString = "SELECT MAX(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE COUNTING = '' AND ID_AB_FILE = '" + ID_AB_FILE + "'";
            conn.ExecuteQuery();

            string ID_PARENT_MAX = conn.GetFieldValue(0, 0);

            conn.QueryString = "SELECT MIN(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE COUNTING = '' AND ID_AB_FILE = '" + ID_AB_FILE + "'";
            conn.ExecuteQuery();

            string ID_PARENT_MIN = conn.GetFieldValue(0, 0);

            try
            {
                if (ID_PARENT_MAX != ID_PARENT_MIN)
                {
                    conn.QueryString = "SELECT QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH FROM AB_CHILDREN_PRK WHERE ID_PARENT = " + ID_PARENT_MAX + " ORDER BY ID_CHILDREN ASC";
                    conn.ExecuteQuery();
                    DataTable dt3 = conn.GetDataTable().Copy();

                    for (int i = 0; i < dt3.Rows.Count; i++)
                    {
                        string QUERYC = dt3.Rows[i]["QUERY_CHILD"].ToString().Trim();
                        string ID_CHILDREN = dt3.Rows[i]["ID_CHILDREN"].ToString().Trim();
                        string DATA_DEC = dt3.Rows[i]["DATA_DEC"].ToString().Trim();
                        string DATA_JENIS = dt3.Rows[i]["DATA_JENIS"].ToString().Trim();
                        string DATA_LENGTH = dt3.Rows[i]["DATA_LENGTH"].ToString().Trim();

                        QUERYC = QUERYC.Replace("#AP_REGNO", "'" + regno + "'");
                        conn.QueryString = QUERYC;
                        conn.ExecuteQuery();

                        string content = conn.GetFieldValue("CONTENT");

                        if (DATA_JENIS == "N")
                        {
                            if (DATA_DEC != "")
                            {
                                int DEC = int.Parse(DATA_DEC);
                                int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                                string[] contents = content.Split(new char[] { ',' });

                                if (contents.Length == 1)
                                {
                                    if (DEC > 0)
                                    {
                                        content = contents[0].PadLeft(SELISIH, '0');
                                        content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                    }
                                    else
                                    {
                                        content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                    }
                                }
                                else
                                {
                                    string content1 = contents[0];
                                    string content2 = contents[1];

                                    content1 = content1.Replace(".", "");
                                    content1 = content1.PadLeft(SELISIH, '0');
                                    content2 = content2.PadRight(DEC, '0');

                                    content = content1 + content2;
                                }
                            }
                            else
                            {
                                content = content.Replace(".", "");
                                content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                            }
                        }
                        else if (DATA_JENIS == "D")
                        {
                            try
                            {
                                DateTime dt = DateTime.Parse(content);
                                string date = String.Format("{0:dd-MM-yyyy}", dt);
                                content = date.Replace("-", "");
                            }
                            catch
                            {
                                content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                            }
                        }

                        stream.Write(content + Environment.NewLine);
                    }
                }
            }
            catch
            {

            }
        }

        public void FillHeaderPRK(StreamWriter stream, string regno, string ID_AB_FILE)
        {
            conn.QueryString = "SELECT MIN(ID_PARENT) as ID_PARENT FROM AB_PARENT_LOOPER_PRK WHERE COUNTING = '' AND ID_AB_FILE = '" + ID_AB_FILE + "'";
            conn.ExecuteQuery();

            try
            {
                string ID_PARENT = conn.GetFieldValue(0, 0);

                conn.QueryString = "SELECT QUERY_CHILD, ID_CHILDREN, DATA_DEC, DATA_JENIS, DATA_LENGTH FROM AB_CHILDREN_PRK WHERE ID_PARENT = " + ID_PARENT + " ORDER BY ID_CHILDREN ASC";
                conn.ExecuteQuery();
                DataTable dt3 = conn.GetDataTable().Copy();

                for (int i = 0; i < dt3.Rows.Count; i++)
                {
                    string QUERYC = dt3.Rows[i]["QUERY_CHILD"].ToString().Trim();
                    string ID_CHILDREN = dt3.Rows[i]["ID_CHILDREN"].ToString().Trim();
                    string DATA_DEC = dt3.Rows[i]["DATA_DEC"].ToString().Trim();
                    string DATA_JENIS = dt3.Rows[i]["DATA_JENIS"].ToString().Trim();
                    string DATA_LENGTH = dt3.Rows[i]["DATA_LENGTH"].ToString().Trim();

                    QUERYC = QUERYC.Replace("#AP_REGNO", "'" + regno + "'");
                    conn.QueryString = QUERYC;
                    conn.ExecuteQuery();

                    string content = conn.GetFieldValue("CONTENT");

                    if (DATA_JENIS == "N")
                    {
                        if (DATA_DEC != "")
                        {
                            int DEC = int.Parse(DATA_DEC);
                            int SELISIH = int.Parse(DATA_LENGTH) - DEC;

                            string[] contents = content.Split(new char[] { ',' });

                            if (contents.Length == 1)
                            {
                                if (DEC > 0)
                                {
                                    content = contents[0].PadLeft(SELISIH, '0');
                                    content = content.PadRight(int.Parse(DATA_LENGTH), '0');
                                }
                                else
                                {
                                    content = contents[0].PadLeft(int.Parse(DATA_LENGTH), '0');
                                }
                            }
                            else
                            {
                                string content1 = contents[0];
                                string content2 = contents[1];

                                content1 = content1.Replace(".", "");
                                content1 = content1.PadLeft(SELISIH, '0');
                                content2 = content2.PadRight(DEC, '0');

                                content = content1 + content2;
                            }
                        }
                        else
                        {
                            content = content.Replace(".", "");
                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                        }
                    }
                    else if (DATA_JENIS == "D")
                    {
                        try
                        {
                            DateTime dt = DateTime.Parse(content);
                            string date = String.Format("{0:dd-MM-yyyy}", dt);
                            content = date.Replace("-", "");
                        }
                        catch
                        {
                            content = content.PadLeft(int.Parse(DATA_LENGTH), '0');
                        }
                    }

                    stream.Write(content + Environment.NewLine);
                }
            }
            catch
            {

            }
        }
    }
}
