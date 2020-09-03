using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.ServiceModel;
using System.Web;
using System.Configuration;
using System.Data;
using DMS.DBConnection;
using DMS.CuBESCore;
using Microsoft.VisualBasic;
using System.Drawing;
using System.IO;

namespace SMEExportImportService
{
    [ServiceBehavior(
        ConcurrencyMode = ConcurrencyMode.Single,
        InstanceContextMode = InstanceContextMode.PerCall/*,
        ReleaseServiceInstanceOnTransactionComplete = true*/
      )]
    class ExportWord : IWord
    {
        protected Tools tool = new Tools();

        private void ReplaceBookmarkText(Microsoft.Office.Interop.Word.Document doc, string bookmarkName, string text)
        {
            if (doc.Bookmarks.Exists(bookmarkName))
            {
                Object name = bookmarkName;
                Microsoft.Office.Interop.Word.Range range =
                    doc.Bookmarks.get_Item(ref name).Range;

                range.Text = text;
                object newRange = range;
                doc.Bookmarks.Add(bookmarkName, ref newRange);
            }
        }

        private string GetBookmarkText(Microsoft.Office.Interop.Word.Document doc, string bookmarkName)
        {
            string thetext = "";

            if (doc.Bookmarks.Exists(bookmarkName))
            {
                Object name = bookmarkName;
                Microsoft.Office.Interop.Word.Range range =
                    doc.Bookmarks.get_Item(ref name).Range;

                thetext = range.Text;
            }

            return thetext;
        }

        private string myMoneyFormat_noDec(string str)
        {
            if ((str.Trim() == "") || (str.Trim() == "&nbsp;"))
            {
                return Strings.FormatNumber(0, 0, TriState.UseDefault, TriState.UseDefault, TriState.UseDefault);
            }
            else
            {
                return Strings.FormatNumber(str, 0, TriState.UseDefault, TriState.UseDefault, TriState.UseDefault);
            }
        }

        private string formatMoney_ind(string a)
        {
            string b, c, d;																	//a = 1,230.00
            c = Strings.Replace(myMoneyFormat_noDec(a), ".", ";", 1, -1, CompareMethod.Binary);	//c = 1,230;00
            b = Strings.Replace(c, ",", ".", 1, -1, CompareMethod.Binary);						//b = 1.230;00
            d = Strings.Replace(b, ";", ",", 1, -1, CompareMethod.Binary);						//d = 1.230,00

            return d;
            //return myMoneyFormat_noDec(a);
        }

        protected Connection conn = new Connection(ConfigurationManager.AppSettings["conn"]);

        void IWord.ExportWord(string name)
        {
            ArrayList orgId = new ArrayList();
            ArrayList newId = new ArrayList();

            Process[] oldProcess = Process.GetProcessesByName("WINWORD");
            foreach (Process thisProcess in oldProcess)
                orgId.Add(thisProcess);

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;
            word.Visible = false;
            word.ScreenUpdating = false;

            Process[] newProcess = Process.GetProcessesByName("WINWORD");
            foreach (Process thisProcess in newProcess)
                newId.Add(thisProcess);

            string fileName = @"C:\TEMPLATE\New Microsoft Office Word Document.docx";
            string saveas = @"C:\Template\TES-" + name + ".docx";
            Document doc = word.Documents.Open(fileName, ref oMissing,
                    false, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            doc.Activate();

            try
            {
                if (doc != null)
                {
                    ReplaceBookmarkText(doc, "pras", "Found !");
                    ReplaceBookmarkText(doc, "pras1", "Found Again !");
                    ReplaceBookmarkText(doc, "pras2", "Found Agaiiiiin :D :D :D !");
                }
                else
                {
                   
                }
            }
            catch (Exception f)
            {
                string a = f.Message;
            }
            finally
            {
                doc.SaveAs(saveas, ref oMissing, false, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                if (doc != null)
                { 
                    //doc.Close(true, oMissing, oMissing); 
                    ((Microsoft.Office.Interop.Word._Document)doc).Close(true, oMissing, oMissing);
                }

                if (word != null)
                {
                    ((Microsoft.Office.Interop.Word._Application)word).Quit(true, oMissing, oMissing); 
                }
            }

            try
            {

                // Killing Proses after Export
                for (int x = 0; x < newId.Count; x++)
                {
                    Process xnewId = (Process)newId[x];

                    bool bSameId = false;
                    for (int z = 0; z < orgId.Count; z++)
                    {
                        Process xoldId = (Process)orgId[z];

                        if (xnewId.Id == xoldId.Id)
                        {
                            bSameId = true;
                            break;
                        }
                    }

                    if (bSameId)
                    {
                        try
                        {
                            xnewId.Kill();
                        }
                        catch
                        {
                            continue;
                        }
                    }

                } // end x		
            }
            catch
            {
                
            }
        }

        string IWord.DocumentExportASCXCreateWord(string templateid, string regno, string userid)
        {
            
            string templatefilename = "",
                outputfilename = "",
                templatepath = "",
                outputpath = "";
            string returnmsg = string.Empty;
            int writeitem = 0;
            bool savestatus;
            try
            {
                //string fileIn = string.Empty;
                //string fileOut = string.Empty;
                object fileIn;
                object fileOut;
                System.Data.DataTable dt1;
                System.Data.DataTable dt2;
                System.Data.DataTable dt3;

                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                object oMissingObject = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document wordDoc = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                wordApp = new Application();
                wordApp.Visible = false;

                //Collecting Existing Word in Taskbar
                Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);

                //Get Export Properties
                conn.QueryString = "SELECT TOP 1 * FROM VW_DOCEXPORT_PARAMETER WHERE TEMPLATE_ID = '" + templateid + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    templatefilename = conn.GetFieldValue("TEMPLATE_FILENAME");
                    outputfilename = conn.GetFieldValue("UPLOAD_FILEFORMAT").Replace("#REGNO$", regno).Replace("#USERID$", userid) + ".DOCX";
                    /*templatepath = HttpContext.Current.Server.MapPath(conn.GetFieldValue("TEMPLATE_PATH").Trim());
                    outputpath = HttpContext.Current.Server.MapPath(conn.GetFieldValue("UPLOAD_PATH").Trim());*/

                    //string abc = ConfigurationManager.AppSettings["serverPath"];

                    templatepath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("TEMPLATE_PATH");
                    outputpath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("UPLOAD_PATH");

                    templatepath = templatepath.Replace("..", "");
                    outputpath = outputpath.Replace("..", "");

                    try
                    {
                        //Collectiong Existing Word in Taskbar
                        Process[] newProcess;
                        try
                        {
                            newProcess = Process.GetProcessesByName("WINWORD");

                            foreach (Process thisProcess in newProcess)
                            {
                                try
                                {
                                    newId.Add(thisProcess);
                                }
                                catch (Exception x)
                                {
                                    returnmsg = x.Message;
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            returnmsg = e.Message;
                        }

                        //Save process into database
                        //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);

                        //fileIn = @templatepath + templatefilename;
                        fileIn = templatepath + templatefilename.Replace(".DOCX", ".docx");
                        try
                        {

                            wordDoc = wordApp.Documents.Open(ref fileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                                ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            //wordDoc.Activate();
                            //						new System.Threading.ThreadStart(wordApp.Documents.Open(ref fileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, 
                            //							ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject)).Start();

                        }
                        catch (Exception ex1)
                        {
                            returnmsg = ex1.Message;
                        }

                        try
                        {
                            wordDoc.Activate();
                        }
                        catch (Exception ex1)
                        {
                            returnmsg = ex1.Message;
                        }

                        Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;
                        //Loop for Template Master

                        conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM DOCEXPORT_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                        conn.ExecuteQuery();

                        dt1 = conn.GetDataTable().Copy();


                        if (dt1.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt1.Rows.Count; i++)
                            {
                                string sheetid = dt1.Rows[i][0].ToString().Trim();
                                string sheetseq = dt1.Rows[i][1].ToString().Trim();
                                string proc = dt1.Rows[i][2].ToString().Trim();

                                //Query Stored Procedure
                                conn.QueryString = "EXEC " + proc + " '" + regno + "'";
                                conn.ExecuteQuery();
                                dt3 = conn.GetDataTable().Copy();

                                //Loop for Template Detail
                                conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM DOCEXPORT_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                        "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                                conn.ExecuteQuery();
                                dt2 = conn.GetDataTable().Copy();

                                for (int k = 0; k < dt3.Rows.Count; k++)
                                {
                                    for (int j = 0; j < dt2.Rows.Count; j++)
                                    {
                                        string xarr = dt2.Rows[j][0].ToString().Trim(); //indicating "0"=array, "1"=single data
                                        object wbm = dt2.Rows[j][1].ToString().Trim(); //bookmark di wordnya
                                        string dbfield = dt2.Rows[j][2].ToString().Trim();
                                        string cell_value = dt3.Rows[k][dbfield].ToString().Trim();

                                        if (wordBookMark.Exists(wbm.ToString()))
                                        {
                                            //if (xarr == "0") cell_value = cell_value + "\n";
                                            if (xarr == "0")
                                            {
                                                string oldtext = GetBookmarkText(wordDoc, wbm.ToString());
                                                cell_value = oldtext + "\n" + cell_value;
                                            }

                                            ReplaceBookmarkText(wordDoc, wbm.ToString(), cell_value);
                                            /*
                                            Word.Bookmark oBook = wordBookMark.Item(ref wbm);
                                            oBook.Select();
                                            oBook.Range.Text = cell_value;*/

                                            writeitem++;
                                        }
                                    }
                                }
                            }


                            //if (writeitem > 0)
                            //{
                            //Save Word File
                            outputfilename = outputfilename.Replace(".doc", ".docx");
                            fileOut = outputpath + outputfilename;
                            //fileOut = @"C:\inetpub\wwwroot\SME\FileUpload\" + outputfilename;
                            wordDoc.SaveAs(ref fileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                                ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));

                            savestatus = true;
                            //}
                            //else
                            //{
                            //	savestatus = false;
                            //	returnmsg = "Error in Saving File!!";
                            //	return returnmsg;
                            //}

                            if (savestatus == true)
                            {
                                //Save to Table
                                conn.QueryString = "EXEC DOCEXPORT_SAVE '1', '" +
                                    templateid + "', '" +
                                    regno + "', '" +
                                    userid + "', '" +
                                    outputfilename + "'";
                                conn.ExecuteQuery();

                                //View Upload Files
                                //ViewExportFiles();

                                returnmsg = "Export Success!!";
                            }
                        }
                        else
                        {
                            returnmsg = "Template Procedure Not Yet Defined!!";
                            return returnmsg;
                        }
                    }
                    catch (Exception e)
                    {
                        //Return Fail Message
                        returnmsg = e.Message + "\n" + e.StackTrace;
                    }
                    finally
                    {
                        if (wordDoc != null)
                        {
                            ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            wordDoc = null;
                        }
                        if (wordApp != null)
                        {
                            ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            wordApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        returnmsg = e.Message;
                    }
                }
                else
                {
                    returnmsg = "Export Parameter Not Yet Defined!!";
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return returnmsg;
        }

        string IWord.DocumentExportASCXCreateWordPk(string templateid, string regno, string seq, string userid)
        {
            string templatefilename = "",
                    outputfilename = "",
                    templatepath = "",
                    outputpath = "";
            string returnmsg = string.Empty;
            int writeitem = 0;
            bool savestatus;
            try
            {
                //string fileIn = string.Empty;
                //string fileOut = string.Empty;
                object fileIn;
                object fileOut;
                System.Data.DataTable dt1;
                System.Data.DataTable dt2;
                System.Data.DataTable dt3;

                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                object oMissingObject = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document wordDoc = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                wordApp = new Application();
                wordApp.Visible = false;

                //Collecting Existing Word in Taskbar
                Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);

                //Get Export Properties
                conn.QueryString = "SELECT TOP 1 * FROM VW_DOCEXPORT_PARAMETER WHERE TEMPLATE_ID = '" + templateid + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    templatefilename = conn.GetFieldValue("TEMPLATE_FILENAME");
                    outputfilename = conn.GetFieldValue("UPLOAD_FILEFORMAT").Replace("#REGNO$", regno).Replace("#PRODUCT$", templateid).Replace("#SEQ$", seq).Replace("#USERID$", userid) + ".DOCX";
                    /*templatepath = HttpContext.Current.Server.MapPath(conn.GetFieldValue("TEMPLATE_PATH").Trim());
                    outputpath = HttpContext.Current.Server.MapPath(conn.GetFieldValue("UPLOAD_PATH").Trim());*/

                    //string abc = ConfigurationManager.AppSettings["serverPath"];

                    templatepath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("TEMPLATE_PATH");
                    outputpath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("UPLOAD_PATH");

                    templatepath = templatepath.Replace("..", "");
                    outputpath = outputpath.Replace("..", "");

                    try
                    {
                        //Collectiong Existing Word in Taskbar
                        Process[] newProcess;
                        try
                        {
                            newProcess = Process.GetProcessesByName("WINWORD");

                            foreach (Process thisProcess in newProcess)
                            {
                                try
                                {
                                    newId.Add(thisProcess);
                                }
                                catch (Exception x)
                                {
                                    returnmsg = x.Message;
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            returnmsg = e.Message;
                        }

                        //Save process into database
                        //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);

                        //fileIn = @templatepath + templatefilename;
                        fileIn = templatepath + templatefilename.Replace(".DOCX", ".docx");
                        try
                        {

                            wordDoc = wordApp.Documents.Open(ref fileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                                ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            //wordDoc.Activate();
                            //						new System.Threading.ThreadStart(wordApp.Documents.Open(ref fileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, 
                            //							ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject)).Start();

                        }
                        catch (Exception ex1)
                        {
                            returnmsg = ex1.Message;
                        }

                        try
                        {
                            wordDoc.Activate();
                        }
                        catch (Exception ex1)
                        {
                            returnmsg = ex1.Message;
                        }

                        Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;
                        //Loop for Template Master

                        conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM DOCEXPORT_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                        conn.ExecuteQuery();

                        dt1 = conn.GetDataTable().Copy();


                        if (dt1.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt1.Rows.Count; i++)
                            {
                                string sheetid = dt1.Rows[i][0].ToString().Trim();
                                string sheetseq = dt1.Rows[i][1].ToString().Trim();
                                string proc = dt1.Rows[i][2].ToString().Trim();

                                //Query Stored Procedure
                                conn.QueryString = "EXEC " + proc + " '" + regno + "','" + seq + "','" + userid + "'";
                                conn.ExecuteQuery();
                                dt3 = conn.GetDataTable().Copy();

                                //Loop for Template Detail
                                conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM DOCEXPORT_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                        "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                                conn.ExecuteQuery();
                                dt2 = conn.GetDataTable().Copy();

                                for (int k = 0; k < dt3.Rows.Count; k++)
                                {
                                    for (int j = 0; j < dt2.Rows.Count; j++)
                                    {
                                        string xarr = dt2.Rows[j][0].ToString().Trim(); //indicating "0"=array, "1"=single data
                                        object wbm = dt2.Rows[j][1].ToString().Trim(); //bookmark di wordnya
                                        string dbfield = dt2.Rows[j][2].ToString().Trim();
                                        string cell_value = dt3.Rows[k][dbfield].ToString().Trim();

                                        if (wordBookMark.Exists(wbm.ToString()))
                                        {
                                            //if (xarr == "0") cell_value = cell_value + "\n";
                                            if (xarr == "0")
                                            {
                                                string oldtext = GetBookmarkText(wordDoc, wbm.ToString());
                                                cell_value = oldtext + "\n" + cell_value;
                                            }

                                            ReplaceBookmarkText(wordDoc, wbm.ToString(), cell_value);
                                            /*
                                            Word.Bookmark oBook = wordBookMark.Item(ref wbm);
                                            oBook.Select();
                                            oBook.Range.Text = cell_value;*/

                                            writeitem++;
                                        }
                                    }
                                }
                            }

                            //if (writeitem > 0)
                            //{
                            //Save Word File
                            outputfilename = outputfilename.Replace(".doc", ".docx");
                            fileOut = outputpath + outputfilename;
                            //fileOut = @"C:\inetpub\wwwroot\SME\FileUpload\" + outputfilename;
                            wordDoc.SaveAs(ref fileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                                ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));

                            savestatus = true;
                            //}
                            //else
                            //{
                            //	savestatus = false;
                            //	returnmsg = "Error in Saving File!!";
                            //	return returnmsg;
                            //}

                            if (savestatus == true)
                            {
                                //Save to Table
                                conn.QueryString = "EXEC DOCEXPORT_SAVE '1', '" +
                                                    templateid + "', '" +
                                                    regno + "', '" +
                                                    userid + "', '" +
                                                    outputfilename + "'";
                                conn.ExecuteQuery();

                                //View Upload Files
                                //ViewExportFiles();

                                returnmsg = "Export Success!!";
                            }
                        }
                        else
                        {
                            returnmsg = "Template Procedure Not Yet Defined!!";
                            return returnmsg;
                        }
                    }
                    catch (Exception e)
                    {
                        //Return Fail Message
                        returnmsg = e.Message + "\n" + e.StackTrace;
                    }
                    finally
                    {
                        if (wordDoc != null)
                        {
                            ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            wordDoc = null;
                        }
                        if (wordApp != null)
                        {
                            ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            wordApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        returnmsg = e.Message;
                    }
                }
                else
                {
                    returnmsg = "Export Parameter Not Yet Defined!!";
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return returnmsg;
        }

        string IWord.DocumentExportASCXCreateExcel(string templateid, string regno, string userid)
        {
            string templatefilename = "",
                outputfilename = "",
                templatepath = "",
                outputpath = "";
            string returnmsg = string.Empty;
            int writeitem = 0;
            bool savestatus;
            try
            {
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                System.Data.DataTable dt1;
                System.Data.DataTable dt2;
                System.Data.DataTable dt3;

                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                //Collecting Existing Excel in Taskbar
                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);

                //Get Export Properties
                conn.QueryString = "SELECT TOP 1 * FROM VW_DOCEXPORT_PARAMETER WHERE TEMPLATE_ID = '" + templateid + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    templatefilename = conn.GetFieldValue("TEMPLATE_FILENAME");
                    outputfilename = conn.GetFieldValue("UPLOAD_FILEFORMAT").Replace("#REGNO$", regno).Replace("#USERID$", userid) + ".XLSX";
                    //templatepath = Server.MapPath(conn.GetFieldValue("TEMPLATE_PATH").Trim());
                    //outputpath = Server.MapPath(conn.GetFieldValue("UPLOAD_PATH").Trim());

                    templatepath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("TEMPLATE_PATH");
                    outputpath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("UPLOAD_PATH");

                    templatepath = templatepath.Replace("..", "");
                    outputpath = outputpath.Replace("..", "");

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        //Collectiong Existing Excel in Taskbar
                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        //Save process into database
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = templatepath + templatefilename;
                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;

                        //Loop for Template Master
                        conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM DOCEXPORT_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                        conn.ExecuteQuery();

                        dt1 = conn.GetDataTable().Copy();

                        if (dt1.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt1.Rows.Count; i++)
                            {
                                string sheetid = dt1.Rows[i][0].ToString().Trim();
                                string sheetseq = dt1.Rows[i][1].ToString().Trim();
                                string proc = dt1.Rows[i][2].ToString().Trim();

                                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheetid);

                                //Query Stored Procedure
                                conn.QueryString = "EXEC " + proc + " '" + regno + "'";
                                conn.ExecuteQuery();
                                dt3 = conn.GetDataTable().Copy();

                                if (dt3.Rows.Count > 0)
                                {
                                    //Loop for Template Detail
                                    conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM DOCEXPORT_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                        "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                                    conn.ExecuteQuery();
                                    dt2 = conn.GetDataTable().Copy();

                                    for (int j = 0; j < dt2.Rows.Count; j++)
                                    {
                                        string xrow = dt2.Rows[j][0].ToString().Trim();
                                        string xcol = dt2.Rows[j][1].ToString().Trim();
                                        string dbfield = dt2.Rows[j][2].ToString().Trim();
                                        string cell_value = dt3.Rows[0][dbfield].ToString().Trim();
                                        string xcell = xcol + xrow;

                                        Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                        if (excelCell != null)
                                        {
                                            excelCell.Value2 = cell_value;
                                            writeitem++;
                                        }
                                    }
                                }
                                else
                                {
                                    //returnmsg = "Query Has No Row!!";
                                    //return returnmsg;
                                }
                            }


                            //if (writeitem > 0)
                            //{
                            //Save Excel File
                            fileOut = outputpath + outputfilename;
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            savestatus = true;
                            //}
                            //else
                            //{
                            //	savestatus = false;
                            //	returnmsg = "Error in Saving File!!";
                            //	return returnmsg;
                            //}

                            if (savestatus == true)
                            {
                                //Save to Table
                                conn.QueryString = "EXEC DOCEXPORT_SAVE '1', '" +
                                    templateid + "', '" +
                                    regno + "', '" +
                                    userid + "', '" +
                                    outputfilename + "'";
                                conn.ExecuteQuery();

                                //View Upload Files
                                //ViewExportFiles();

                                returnmsg = "Export Success!!";
                            }
                        }
                        else
                        {
                            returnmsg = "Template Procedure Not Yet Defined!!";
                            return returnmsg;
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");

                        //Return Fail Message
                        returnmsg = e.Message + "\n" + e.StackTrace;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (!bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch { }
                }
                else
                {
                    returnmsg = "Export Parameter Not Yet Defined!!";
                    return returnmsg;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return returnmsg;
        }

        string IWord.DocumentExportASCXCreateExcel2(string templateid, string regno, string userid)
        {
            string templatefilename = "",
                outputfilename = "",
                templatepath = "",
                outputpath = "";
            string returnmsg = string.Empty;
            int writeitem = 0;
            bool savestatus;

            try
            {

                string fileIn = string.Empty;
                string fileOut = string.Empty;
                System.Data.DataTable dt1;
                System.Data.DataTable dt2;
                System.Data.DataTable dt3;

                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                //Collecting Existing Excel in Taskbar
                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);

                //Get Export Properties
                conn.QueryString = "SELECT TOP 1 * FROM VW_DOCEXPORT_PARAMETER WHERE TEMPLATE_ID = '" + templateid + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    templatefilename = conn.GetFieldValue("TEMPLATE_FILENAME");
                    outputfilename = conn.GetFieldValue("UPLOAD_FILEFORMAT").Replace("#REGNO$", regno).Replace("#USERID$", userid) + ".XLSX";
                    //templatepath = Server.MapPath(conn.GetFieldValue("TEMPLATE_PATH").Trim());
                    //outputpath = Server.MapPath(conn.GetFieldValue("UPLOAD_PATH").Trim());
                    templatepath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("TEMPLATE_PATH");
                    outputpath = ConfigurationManager.AppSettings["serverPath"] + conn.GetFieldValue("UPLOAD_PATH");

                    templatepath = templatepath.Replace("..", "");
                    outputpath = outputpath.Replace("..", "");

                    conn.QueryString = "INSERT INTO [DEBUG_SP] ([SP_NAME],[SP_ARG],[DBG_DATE]) VALUES ('','" + templatepath + "','" + DateAndTime.Now + "')";
                    conn.ExecuteQuery();

                    conn.QueryString = "INSERT INTO [DEBUG_SP] ([SP_NAME],[SP_ARG],[DBG_DATE]) VALUES ('','" + outputpath + "','" + DateAndTime.Now + "')";
                    conn.ExecuteQuery();

                    //try
                    //{
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        //Collectiong Existing Excel in Taskbar
                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        //Save process into database
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = templatepath + templatefilename.Replace(".XLSX", ".XLSX");
                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;

                        //Loop for Template Master
                        conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM DOCEXPORT_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                        conn.ExecuteQuery();

                        dt1 = conn.GetDataTable().Copy();

                        if (dt1.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt1.Rows.Count; i++)
                            {
                                string sheetid = dt1.Rows[i][0].ToString().Trim();
                                string sheetseq = dt1.Rows[i][1].ToString().Trim();
                                string proc = dt1.Rows[i][2].ToString().Trim();

                                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheetid);

                                //Query Stored Procedure
                                conn.QueryString = "EXEC " + proc + " '" + regno + "'";
                                conn.ExecuteQuery();
                                dt3 = conn.GetDataTable().Copy();

                                if (dt3.Rows.Count > 0)
                                {
                                    //Loop for Template Detail
                                    conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM DOCEXPORT_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                        "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                                    conn.ExecuteQuery();
                                    dt2 = conn.GetDataTable().Copy();

                                    for (int k = 0; k < dt3.Rows.Count; k++)
                                    {
                                        for (int j = 0; j < dt2.Rows.Count; j++)
                                        {
                                            int irow;
                                            try { irow = int.Parse(dt2.Rows[j][0].ToString().Trim()) + k; }
                                            catch { irow = 1; }
                                            string xrow = irow.ToString().Trim();
                                            string xcol = dt2.Rows[j][1].ToString().Trim();
                                            string dbfield = dt2.Rows[j][2].ToString().Trim();
                                            string cell_value = dt3.Rows[k][dbfield].ToString().Trim();
                                            string xcell = xcol + xrow;

                                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                            if (excelCell != null)
                                            {
                                                excelCell.Value2 = cell_value;
                                                writeitem++;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    returnmsg = "Query Has No Row!!";
                                    //return returnmsg;
                                }
                            }


                            //if (writeitem > 0)
                            //{
                            //Save Excel File
                            fileOut = outputpath + outputfilename.Replace(".XLSX", ".XLSX");
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            savestatus = true;
                            //}
                            //else
                            //{
                            //	savestatus = false;
                            //	returnmsg = "Error in Saving File!!";
                            //	return returnmsg;
                            //}

                            if (savestatus == true)
                            {
                                //Save to Table
                                conn.QueryString = "EXEC DOCEXPORT_SAVE '1', '" +
                                    templateid + "', '" +
                                    regno + "', '" +
                                    userid + "', '" +
                                    outputfilename + "'";
                                conn.ExecuteQuery();

                                //View Upload Files
                                //ViewExportFiles();

                                returnmsg = "Export Success!!";
                            }
                        }
                        else
                        {
                            returnmsg = "Template Procedure Not Yet Defined!!";
                            return returnmsg;
                        }
                    //}
                    //catch (Exception e)
                    //{
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");

                        //Return Fail Message
                        //returnmsg = e.Message + "\n" + e.StackTrace;
                    //}
                    //finally
                    //{
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    //}

                    //try
                    //{
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    //}
                    //catch { }
                }
                else
                {
                    returnmsg = "Export Parameter Not Yet Defined!!";
                    //return returnmsg;
                }
            }
            catch (Exception ex)
            {
                conn.QueryString = "INSERT INTO [DEBUG_SP] ([SP_NAME],[SP_ARG],[DBG_DATE]) VALUES ('','" + ex.ToString() + "','" + DateAndTime.Now + "')";
                conn.ExecuteQuery();
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return returnmsg;
        }

        string IWord.DocumentUploadASCXReadExcel(string filename, string templateid, string regno)
        {
            string resultmsg = "";

            /*System.Web.UI.WebControls.Label LBL_STATUSREPORT = (System.Web.UI.WebControls.Label)thisControl.FindControl("LBL_STATUSREPORT");
            System.Web.UI.WebControls.Label LBL_STATUS = (System.Web.UI.WebControls.Label)thisControl.FindControl("LBL_STATUS");*/
            try
            {

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);

                System.Data.DataTable dt1, dt2;

                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);

                    //Save process into database
                    //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);
                    filename = filename.Replace(".XLSX", ".XLSX");
                    excelWorkBook = excelApp.Workbooks.Open(filename,
                        0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        false, false, 0, true);

                    excelSheet = excelWorkBook.Worksheets;

                    //Loop for Template Master
                    conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM DOCEXPORT_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                    conn.ExecuteQuery();

                    dt1 = conn.GetDataTable().Copy();

                    if (dt1.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            string sheetid = dt1.Rows[i][0].ToString().Trim();
                            string sheetseq = dt1.Rows[i][1].ToString().Trim();
                            string proc = dt1.Rows[i][2].ToString().Trim();

                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheetid);

                            //Loop for Template Detail
                            conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM DOCEXPORT_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                            conn.ExecuteQuery();
                            dt2 = conn.GetDataTable().Copy();
                            int n = dt2.Rows.Count;
                            object[] par;
                            par = new object[n];
                            object[] dttype;
                            dttype = new object[n];
                            
                            if (templateid == "PUNDI_CAS")
                            {
                                if (sheetseq == "1")
                                {
                                    conn.QueryString = "DELETE PUNDI_CAS WHERE AP_REGNO = '" + regno + "'";
                                    conn.ExecuteNonQuery();

                                    int spare = 0;
                                    bool done = false;
                                    while(true)
                                    {
                                        //looping disini
                                        for (int j = 0; j < dt2.Rows.Count; j++)
                                        {
                                            string xrow = (int.Parse(dt2.Rows[j][0].ToString().Trim()) + spare).ToString();
                                            string xcol = dt2.Rows[j][1].ToString().Trim();
                                            string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                            string cell_value;
                                            string xcell = xcol + xrow;

                                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                            try
                                            {
                                                if (excelCell != null)
                                                {
                                                    cell_value = excelCell.Value2.ToString();
                                                    par[j] = (string)cell_value;
                                                    dttype[j] = (string)datatype;
                                                }
                                            }
                                            catch
                                            {
                                                cell_value = "0";
                                                par[j] = "0";
                                                dttype[j] = "N";
                                            }
                                                

                                            //klo nama uda kosong break
                                            if (par[0].ToString() == "0")
                                            {
                                                done = true;
                                                break;
                                            }
                                        }

                                        if (!done)
                                        {
                                            //Construct Query
                                            string query = "EXEC " + proc + " '" + regno + "', ";
                                            for (int k = 0; k < n; k++)
                                            {
                                                if (dttype[k].ToString() == "C")
                                                    query = query + "'" + par[k].ToString() + "'";
                                                else if (dttype[k].ToString() == "N")
                                                    query = query + "" + par[k].ToString() + "";

                                                if (k < n - 1)
                                                    query = query + ", ";
                                            }
                                            //Run Query
                                            conn.QueryString = query;
                                            conn.ExecuteQuery();
                                        }
                                        else
                                        {
                                            break;
                                        }
                                        spare += 74;
                                    }
                                }
                                else if(sheetseq == "2")
                                {
                                    if (dt2.Rows.Count > 0)
                                    {
                                        for (int j = 0; j < dt2.Rows.Count; j++)
                                        {
                                            string xrow = dt2.Rows[j][0].ToString().Trim();
                                            string xcol = dt2.Rows[j][1].ToString().Trim();
                                            string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                            string cell_value;
                                            string xcell = xcol + xrow;

                                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                            if (excelCell != null)
                                            {
                                                cell_value = excelCell.Value2.ToString();
                                                par[j] = (string)cell_value;
                                                dttype[j] = (string)datatype;
                                            }
                                        }

                                        //Construct Query
                                        string query = "EXEC " + proc + " '" + regno + "', ";
                                        for (int k = 0; k < n; k++)
                                        {
                                            if (dttype[k].ToString() == "C")
                                                query = query + "'" + par[k].ToString() + "'";
                                            else if (dttype[k].ToString() == "N")
                                                query = query + "" + par[k].ToString() + "";

                                            if (k < n - 1)
                                                query = query + ", ";
                                        }

                                        //Run Query
                                        conn.QueryString = query;
                                        conn.ExecuteQuery();

                                        //Show Success Message
                                        /*LBL_STATUS.ForeColor = Color.Green;
                                        LBL_STATUSREPORT.ForeColor = Color.Green;
                                        LBL_STATUS.Text = "Upload Sucessful! Insert Result Sucessful!";
                                        LBL_STATUSREPORT.Text = "";*/
                                        resultmsg = "success";
                                    }
                                }
                            }
                            else
                            {
                                if (dt2.Rows.Count > 0)
                                {
                                    for (int j = 0; j < dt2.Rows.Count; j++)
                                    {
                                        string xrow = dt2.Rows[j][0].ToString().Trim();
                                        string xcol = dt2.Rows[j][1].ToString().Trim();
                                        string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                        string cell_value;
                                        string xcell = xcol + xrow;

                                        Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                        if (excelCell != null)
                                        {
                                            cell_value = excelCell.Value2.ToString();
                                            par[j] = (string)cell_value;
                                            dttype[j] = (string)datatype;
                                        }
                                    }

                                    //Construct Query
                                    string query = "EXEC " + proc + " '" + regno + "', ";
                                    for (int k = 0; k < n; k++)
                                    {
                                        if (dttype[k].ToString() == "C")
                                            query = query + "'" + par[k].ToString() + "'";
                                        else if (dttype[k].ToString() == "N")
                                            query = query + "" + par[k].ToString() + "";

                                        if (k < n - 1)
                                            query = query + ", ";
                                    }

                                    //Run Query
                                    conn.QueryString = query;
                                    conn.ExecuteQuery();

                                    //Show Success Message
                                    /*LBL_STATUS.ForeColor = Color.Green;
                                    LBL_STATUSREPORT.ForeColor = Color.Green;
                                    LBL_STATUS.Text = "Upload Sucessful! Insert Result Sucessful!";
                                    LBL_STATUSREPORT.Text = "";*/
                                    resultmsg = "success";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                    /*LBL_STATUS.ForeColor = Color.Red;
                    LBL_STATUSREPORT.ForeColor = Color.Red;
                    LBL_STATUS.Text = "Upload Failed!";
                    LBL_STATUSREPORT.Text = ex.Message + "\n" + ex.StackTrace;*/

                    //Response.Write("<!--" + ex.Message + "\n" + ex.StackTrace + "-->");
                    resultmsg = ex.Message;
                }
                finally
                {
                    if (excelWorkBook != null)
                    {
                        excelWorkBook.Close(true, filename, null);
                        excelWorkBook = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.Workbooks.Close();
                        excelApp.Application.Quit();
                        excelApp = null;
                    }
                }

                try
                {
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    resultmsg = e.Message;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return resultmsg;
        }

        string IWord.DocumentUploadASCXReadExcelPensiun(string filename, string templateid, string regno)
        {
            string resultmsg = "";

            /*System.Web.UI.WebControls.Label LBL_STATUSREPORT = (System.Web.UI.WebControls.Label)thisControl.FindControl("LBL_STATUSREPORT");
            System.Web.UI.WebControls.Label LBL_STATUS = (System.Web.UI.WebControls.Label)thisControl.FindControl("LBL_STATUS");*/
            try
            {

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);

                System.Data.DataTable dt1, dt2;

                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);

                    //Save process into database
                    //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);
                    filename = filename.Replace(".XLSX", ".XLSX");
                    excelWorkBook = excelApp.Workbooks.Open(filename,
                        0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        false, false, 0, true);

                    excelSheet = excelWorkBook.Worksheets;

                    //Loop for Template Master
                    conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM DOCEXPORT_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                    conn.ExecuteQuery();

                    dt1 = conn.GetDataTable().Copy();

                    if (dt1.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            string sheetid = dt1.Rows[i][0].ToString().Trim();
                            string sheetseq = dt1.Rows[i][1].ToString().Trim();
                            string proc = dt1.Rows[i][2].ToString().Trim();

                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheetid);

                            //Loop for Template Detail
                            conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM DOCEXPORT_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                            conn.ExecuteQuery();
                            dt2 = conn.GetDataTable().Copy();
                            int n = dt2.Rows.Count;
                            object[] par;
                            par = new object[n];
                            object[] dttype;
                            dttype = new object[n];

                            if (templateid == "PENSIUN_CALC")
                            {
                                if (sheetseq == "1")
                                {
                                    conn.QueryString = "DELETE PENSIUN_CALC WHERE AP_REGNO = '" + regno + "'";
                                    conn.ExecuteNonQuery();

                                    int spare = 0;
                                    bool done = false;
                                    while (true)
                                    {
                                        //looping disini
                                        for (int j = 0; j < dt2.Rows.Count; j++)
                                        {
                                            string xrow = (int.Parse(dt2.Rows[j][0].ToString().Trim()) + spare).ToString();
                                            string xcol = dt2.Rows[j][1].ToString().Trim();
                                            string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                            string cell_value;
                                            string xcell = xcol + xrow;

                                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                            try
                                            {
                                                if (excelCell != null)
                                                {
                                                    cell_value = excelCell.Value2.ToString();
                                                    par[j] = (string)cell_value;
                                                    dttype[j] = (string)datatype;
                                                }
                                            }
                                            catch
                                            {
                                                cell_value = "0";
                                                par[j] = "0";
                                                dttype[j] = "N";
                                            }


                                            //klo nama uda kosong break
                                            if (par[0].ToString() == "0")
                                            {
                                                done = true;
                                                break;
                                            }
                                        }

                                        if (!done)
                                        {
                                            //Construct Query
                                            string query = "EXEC " + proc + " '" + regno + "', ";
                                            for (int k = 0; k < n; k++)
                                            {
                                                if (dttype[k].ToString() == "C")
                                                    query = query + "'" + par[k].ToString() + "'";
                                                else if (dttype[k].ToString() == "N")
                                                    query = query + "" + par[k].ToString() + "";

                                                if (k < n - 1)
                                                    query = query + ", ";
                                            }
                                            //Run Query
                                            conn.QueryString = query;
                                            conn.ExecuteQuery();
                                        }
                                        else
                                        {
                                            break;
                                        }
                                        spare += 74;
                                    }
                                }
                                else if (sheetseq == "2")
                                {
                                    if (dt2.Rows.Count > 0)
                                    {
                                        for (int j = 0; j < dt2.Rows.Count; j++)
                                        {
                                            string xrow = dt2.Rows[j][0].ToString().Trim();
                                            string xcol = dt2.Rows[j][1].ToString().Trim();
                                            string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                            string cell_value;
                                            string xcell = xcol + xrow;

                                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                            if (excelCell != null)
                                            {
                                                cell_value = excelCell.Value2.ToString();
                                                par[j] = (string)cell_value;
                                                dttype[j] = (string)datatype;
                                            }
                                        }

                                        //Construct Query
                                        string query = "EXEC " + proc + " '" + regno + "', ";
                                        for (int k = 0; k < n; k++)
                                        {
                                            if (dttype[k].ToString() == "C")
                                                query = query + "'" + par[k].ToString() + "'";
                                            else if (dttype[k].ToString() == "N")
                                                query = query + "" + par[k].ToString() + "";

                                            if (k < n - 1)
                                                query = query + ", ";
                                        }

                                        //Run Query
                                        conn.QueryString = query;
                                        conn.ExecuteQuery();

                                        //Show Success Message
                                        /*LBL_STATUS.ForeColor = Color.Green;
                                        LBL_STATUSREPORT.ForeColor = Color.Green;
                                        LBL_STATUS.Text = "Upload Sucessful! Insert Result Sucessful!";
                                        LBL_STATUSREPORT.Text = "";*/
                                        resultmsg = "success";
                                    }
                                }
                            }
                            else
                            {
                                if (dt2.Rows.Count > 0)
                                {
                                    for (int j = 0; j < dt2.Rows.Count; j++)
                                    {
                                        string xrow = dt2.Rows[j][0].ToString().Trim();
                                        string xcol = dt2.Rows[j][1].ToString().Trim();
                                        string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                        string cell_value;
                                        string xcell = xcol + xrow;

                                        Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                        if (excelCell != null)
                                        {
                                            cell_value = excelCell.Value2.ToString();
                                            par[j] = (string)cell_value;
                                            dttype[j] = (string)datatype;
                                        }
                                    }

                                    //Construct Query
                                    string query = "EXEC " + proc + " '" + regno + "', ";
                                    for (int k = 0; k < n; k++)
                                    {
                                        if (dttype[k].ToString() == "C")
                                            query = query + "'" + par[k].ToString() + "'";
                                        else if (dttype[k].ToString() == "N")
                                            query = query + "" + par[k].ToString() + "";

                                        if (k < n - 1)
                                            query = query + ", ";
                                    }

                                    //Run Query
                                    conn.QueryString = query;
                                    conn.ExecuteQuery();

                                    //Show Success Message
                                    /*LBL_STATUS.ForeColor = Color.Green;
                                    LBL_STATUSREPORT.ForeColor = Color.Green;
                                    LBL_STATUS.Text = "Upload Sucessful! Insert Result Sucessful!";
                                    LBL_STATUSREPORT.Text = "";*/
                                    resultmsg = "success";
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {

                    /*LBL_STATUS.ForeColor = Color.Red;
                    LBL_STATUSREPORT.ForeColor = Color.Red;
                    LBL_STATUS.Text = "Upload Failed!";
                    LBL_STATUSREPORT.Text = ex.Message + "\n" + ex.StackTrace;*/

                    //Response.Write("<!--" + ex.Message + "\n" + ex.StackTrace + "-->");
                    resultmsg = ex.Message;
                }
                finally
                {
                    if (excelWorkBook != null)
                    {
                        excelWorkBook.Close(true, filename, null);
                        excelWorkBook = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.Workbooks.Close();
                        excelApp.Application.Quit();
                        excelApp = null;
                    }
                }

                try
                {
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    resultmsg = e.Message;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return resultmsg;
        }

        string IWord.Neraca_KMK_KI_SMALLASPXviewExcel(string dir, string regno, string userid, out Dictionary<string, string> results)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            string vPath;
            string resultmsg = "";

            try
            {
                conn.QueryString = "select xls_dir+''+fu_filename as filexls from CA_FILEUPLOADXL where fu_filename = '" + regno + "-" + userid + "-" + dir + "'";
                conn.ExecuteQuery();
                vPath = conn.GetFieldValue("filexls");

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;

                /////////////////////////////////
                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();
                /////////////////////////////////

                /////////////////////////////////////////////////////////////////
                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);
                ////////////////////////////////////////////////////////////////

                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    ////////////////////////////////////////////////////////////////
                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);
                    ////////////////////////////////////////////////////////////////

                    /// Save process into database
                    /// 					
                    //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);
                    vPath = vPath.Replace(".XLSX", "XLSX");
                    excelWorkbook = excelApp.Workbooks.Open(vPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
                    Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
                    string currentSheet = "LOS";
                    Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet);
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk date periode, number of months lihat excel !!!!!!!!!!!
                    for (int i = 66; i < 69; i++)
                    {
                        for (int j = 1; j < 3; j++)
                        {
                            string vtmp = ((char)i).ToString() + j; //i=66 diconvert ke ascci jd huruf B, di concat dgn j hasilnya B1,B2,C1,C2
                            Microsoft.Office.Interop.Excel.Range excelB2 = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vtmp, vtmp);

                            //System.Web.UI.WebControls.TextBox TXT_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vtmp);
                            //System.Web.UI.WebControls.TextBox TXT_TGL_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_TGL_" + vtmp);
                            //System.Web.UI.WebControls.DropDownList DDL_BLN_ = (System.Web.UI.WebControls.DropDownList)thisPage.FindControl("DDL_BLN_" + vtmp);
                            //System.Web.UI.WebControls.TextBox TXT_YEAR_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_YEAR_" + vtmp);

                            string TXT_ = "TXT_" + vtmp;
                            string TXT_TGL_ = "TXT_TGL_" + vtmp;
                            string DDL_BLN_ = "DDL_BLN_" + vtmp;
                            string TXT_YEAR_ = "TXT_YEAR_" + vtmp;

                            string vals = "";
                            string valsDATE = "";
                            string valsMONTH = "";
                            string valsYEAR = "";

                            if (j % 2 == 0)
                            {
                                try
                                {

                                    //TXT_.Text = excelB2.Value2.ToString();
                                    vals = excelB2.Value2.ToString();
                                }
                                catch
                                {
                                    //TXT_.Text = ""; 
                                    vals = "";
                                }

                                result.Add(TXT_, vals);
                            }

                            else
                            {
                                try
                                {
                                    //TXT_.Text = excelB2.Text.ToString();
                                    vals = excelB2.Text.ToString();

                                    string excdatestr = excelB2.Text.ToString();
                                    int dd = int.Parse(excdatestr.Substring(3, 2)),
                                        mm = int.Parse(excdatestr.Substring(0, 2)),
                                        yy = int.Parse(excdatestr.Substring(6, 2));
                                    if (yy < 50)
                                        yy += 2000;
                                    else
                                        yy += 1900;
                                    DateTime excdate = new DateTime(yy, mm, dd);

                                    valsDATE = dd.ToString();
                                    valsMONTH = mm.ToString();
                                    valsYEAR = yy.ToString();
                                    //GlobalTools.fillDateForm(TXT_TGL_, DDL_BLN_, TXT_YEAR_, excdate);

                                }
                                catch
                                {
                                    /*TXT_.Text = "";
                                    TXT_TGL_.Text = "";
                                    DDL_BLN_.SelectedValue = "";
                                    TXT_YEAR_.Text = "";*/

                                    vals = "";
                                    valsDATE = "";
                                    valsMONTH = "";
                                    valsYEAR = "";
                                }

                                result.Add(TXT_, vals);
                                result.Add(TXT_TGL_, valsDATE);
                                result.Add(DDL_BLN_, valsMONTH);
                                result.Add(TXT_YEAR_, valsYEAR);
                            }
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk cash bank sampe liabilities net worth, lihat excel !!!!!!
                    for (int m = 66; m < 69; m++)
                    {
                        for (int n = 3; n <= 35; n++)
                        {
                            string vRange = ((char)m).ToString() + n;
                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vRange, vRange);
                            //System.Web.UI.WebControls.TextBox TXT_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vRange);

                            string TXT_ = "TXT_" + vRange;
                            string vals = "";
                            string valsDDL = "";

                            if (n == 3 || n == 4)
                            {
                                if (n == 3)
                                {
                                    for (int p = 3; p < 4; p++)
                                    {
                                        string vRg = ((char)m).ToString();
                                        string DDL_ = "DDL_" + vRg + p.ToString();
                                        //System.Web.UI.WebControls.DropDownList DDL_ = (System.Web.UI.WebControls.DropDownList)thisPage.FindControl("DDL_" + vRg + p.ToString());
                                        try
                                        {
                                            /*TXT_.Text = excelCell.Text.ToString();
                                            DDL_.SelectedValue = TXT_.Text;*/
                                            vals = excelCell.Text.ToString();
                                            valsDDL = vals;
                                        }
                                        catch
                                        {
                                            /*TXT_.Text = "";
                                            DDL_.SelectedValue = "-";*/

                                            vals = "";
                                            valsDDL = "-";
                                        }

                                        result.Add(TXT_, vals);
                                        result.Add(DDL_, valsDDL);
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        //TXT_.Text = excelCell.Text.ToString();
                                        vals = excelCell.Text.ToString();
                                    }
                                    catch
                                    {
                                        //TXT_.Text = "";
                                        vals = "";
                                    }

                                    result.Add(TXT_, vals);
                                }
                            }
                            else
                            {
                                try
                                {
                                    //TXT_.Text = formatMoney_ind(excelCell.Value2.ToString()); 
                                    vals = formatMoney_ind(excelCell.Value2.ToString());
                                }
                                catch
                                {
                                    //TXT_.Text = ""; 
                                    vals = "";
                                }

                                result.Add(TXT_, vals);
                            }
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                }
                catch
                {

                }
                finally
                {
                    excelWorkbook.Close(null, null, null);
                    excelApp.Workbooks.Close();
                    excelApp.Application.Quit();
                    excelApp.Quit();
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets); 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    //excelSheets = null; 
                    excelWorkbook = null;
                    excelApp = null;

                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
                results = result;
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return resultmsg;
        }

        string IWord.Neraca_KMK_KI_SMALLASPXviewExcel_LabaRugi(string directori, string regno, string userid, out Dictionary<string, string> results)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            string vPath;
            string resultmsg = "";
            //TODO : Jangan di hardcode !!!
            try
            {
                conn.QueryString = "select xls_dir+''+fu_filename as filexls from CA_FILEUPLOADXL where fu_filename = '" + regno + "-" + userid + "-" + directori + "'";
                conn.ExecuteQuery();
                vPath = conn.GetFieldValue("filexls");


                Microsoft.Office.Interop.Excel.Application excelAppIS = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkbookIS = null;

                /////////////////////////////////
                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();
                /////////////////////////////////

                /////////////////////////////////////////////////////////////////
                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);
                ////////////////////////////////////////////////////////////////

                try
                {
                    // Set the culture and UI culture to the browser's accept language
                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelAppIS = new Microsoft.Office.Interop.Excel.Application();
                    excelAppIS.Visible = false;
                    excelAppIS.DisplayAlerts = false;

                    ////////////////////////////////////////////////////////////////
                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);
                    ////////////////////////////////////////////////////////////////

                    /// Save process into database
                    /// 					
                    //SupportTools.saveProcessExcel(excelAppIS, newId, orgId, conn);
                    vPath = vPath.Replace(".XLSX", ".XLSX");
                    excelWorkbookIS = excelAppIS.Workbooks.Open(vPath,
                        0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        true, false, 0, true);
                    Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbookIS.Worksheets;
                    string currentSheet = "LOS";
                    Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet);
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk date periode, number of months lihat excel !!!!!!!!!!!
                    for (int i = 66; i < 69; i++)
                    {
                        string TXT_ = "";
                        string vals = "";

                        for (int j = 36; j < 38; j++)
                        {
                            string vtmp = ((char)i).ToString() + j; //i=66 diconvert ke ascci jd huruf B, di concat dgn j hasilnya B1,B2,C1,C2
                            Microsoft.Office.Interop.Excel.Range excelB2 = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vtmp, vtmp);
                            //System.Web.UI.WebControls.TextBox TXT_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vtmp);
                            TXT_ = "TXT_" + vtmp;
                            if (j % 2 != 0)
                            {
                                try
                                {
                                    //TXT_.Text = excelB2.Value2.ToString();
                                    vals = excelB2.Value2.ToString();
                                }
                                catch
                                {
                                    //TXT_.Text = "";
                                    vals = "";
                                }
                            }

                            else
                            {
                                try
                                {
                                    //TXT_.Text = excelB2.Text.ToString();
                                    vals = excelB2.Text.ToString(); ;
                                }
                                catch
                                {
                                    //TXT_.Text = "";
                                    vals = "";
                                }
                            }
                            result.Add(TXT_, vals);
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk cash bank sampe liabilities net worth, lihat excel !!!!!!
                    for (int m = 66; m < 69; m++)
                    {
                        string TXT_ = "";
                        string vals = "";

                        for (int n = 38; n <= 55; n++)
                        {
                            string vRange = ((char)m).ToString() + n;
                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vRange, vRange);
                            //System.Web.UI.WebControls.TextBox TXT_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vRange);
                            TXT_ = "TXT_" + vRange;
                            if (n == 38)
                            {
                                try
                                {
                                    //TXT_.Text = excelCell.Value2.ToString();
                                    vals = excelCell.Value2.ToString();
                                }
                                catch
                                {
                                    //TXT_.Text = "";
                                    vals = "";
                                }
                            }
                            else
                            {
                                try
                                {
                                    //TXT_.Text = formatMoney_ind(excelCell.Value2.ToString());
                                    vals = formatMoney_ind(excelCell.Value2.ToString());
                                }
                                catch
                                {
                                    //TXT_.Text = "";
                                    vals = "";
                                }
                            }
                            result.Add(TXT_, vals);
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                }
                catch (Exception e)
                {
                    resultmsg = e.Message;
                }
                finally
                {
                    excelWorkbookIS.Close(null, null, null);
                    excelAppIS.Workbooks.Close();
                    excelAppIS.Application.Quit();
                    excelAppIS.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbookIS);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppIS);
                    excelWorkbookIS = null;
                    excelAppIS = null;

                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }

                results = result;
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return resultmsg;
        }

        string IWord.Neraca_KMK_KI_MediumASPXViewExcel(string dir, string regno, string userid, out Dictionary<string, string> results)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            //TO DO ....
            string vPath;
            string returnMsg = "";

            try
            {
                conn.QueryString = "select xls_dir+''+fu_filename as filexls from CA_FILEUPLOADXL where fu_filename = '" + regno + "-" + userid + "-" + dir + "'";
                conn.ExecuteQuery();
                vPath = conn.GetFieldValue("filexls");

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;

                /////////////////////////////////
                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();
                /////////////////////////////////

                /////////////////////////////////////////////////////////////////
                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);
                ////////////////////////////////////////////////////////////////


                /*try
                {*/

                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    ////////////////////////////////////////////////////////////////
                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);
                    ////////////////////////////////////////////////////////////////

                    /// Save process into database
                    /// 					
                    //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);
                    vPath = vPath.Replace(".XLSX", ".XSLX");
                    excelWorkbook = excelApp.Workbooks.Open(vPath,
                        0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        false, false, 0, true);
                    Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbook.Worksheets;
                    string currentSheet = "LOS";
                    Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet);
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk date periode, number of months lihat excel !!!!!!!!!!!
                    for (int i = 66; i < 70; i++)
                    {
                        string TXT_ = "";
                        string DDL_ = "";
                        string TXT_TGL_ = "";
                        string DDL_BLN_ = "";
                        string TXT_YEAR_ = "";

                        string TXT_VAL = "";
                        string DDL_VAL = "";
                        string TXT_TGL_VAL = "";
                        string DDL_BLN_VAL = "";
                        string TXT_YEAR_VAL = "";

                        for (int j = 1; j < 5; j++)
                        {
                            string vtmp = ((char)i).ToString() + j; //i=66 diconvert ke ascci jd huruf B, di concat dgn j hasilnya B1,B2,C1,C2
                            Microsoft.Office.Interop.Excel.Range excelB2 = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vtmp, vtmp);
                            //System.Web.UI.WebControls.TextBox TXT_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vtmp);
                            TXT_ = "TXT_" + vtmp;

                            //System.Web.UI.WebControls.TextBox TXT_TGL_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_TGL_" + vtmp);
                            //System.Web.UI.WebControls.DropDownList DDL_BLN_ = (System.Web.UI.WebControls.DropDownList)thisPage.FindControl("DDL_BLN_" + vtmp);
                            //System.Web.UI.WebControls.TextBox TXT_YEAR_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_YEAR_" + vtmp);
                            TXT_TGL_ = "TXT_TGL_" + vtmp;
                            DDL_BLN_ = "DDL_BLN_" + vtmp;
                            TXT_YEAR_ = "TXT_YEAR_" + vtmp;

                            if (j != 1)
                            {
                                if (j == 3)
                                {
                                    for (int p = 3; p < 4; p++)
                                    {
                                        string vRg = ((char)i).ToString();
                                        //System.Web.UI.WebControls.TextBox teks = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vRg + p.ToString());
                                        TXT_ = "TXT_" + vRg + p.ToString();
                                        //System.Web.UI.WebControls.DropDownList DDL_ = (System.Web.UI.WebControls.DropDownList)thisPage.FindControl("DDL_" + vRg + p.ToString());
                                        DDL_ = "DDL_" + vRg + p.ToString();
                                        try
                                        {
                                            //teks.Text = excelB2.Value2.ToString();
                                            //DDL_.SelectedValue = teks.Text;

                                            TXT_VAL = excelB2.Value2.ToString();
                                            DDL_VAL = excelB2.Value2.ToString();
                                        }
                                        catch
                                        {
                                            //teks.Text = "";
                                            //DDL_.SelectedValue = "-";

                                            TXT_VAL = "";
                                            DDL_VAL = "";
                                        }

                                        result.Add(TXT_, TXT_VAL);
                                        result.Add(DDL_, DDL_VAL);
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        //TXT_.Text = excelB2.Value2.ToString();
                                        TXT_VAL = excelB2.Value2.ToString();
                                    }
                                    catch
                                    {
                                        //TXT_.Text = "0";
                                        TXT_VAL = "0";
                                    }

                                    result.Add(TXT_, TXT_VAL);
                                }
                            }
                            else
                            {
                                //--------------------------------------------------------------------------------------//
                                try
                                {
                                    DateTime excdatestr = Convert.ToDateTime(tool.FormatDate(excelB2.Text.ToString()));
                                    //GlobalTools.fillDateForm(TXT_TGL_, DDL_BLN_, TXT_YEAR_, excdatestr);
                                    TXT_TGL_VAL = excdatestr.Date.ToString();
                                    DDL_BLN_VAL = excdatestr.Month.ToString();
                                    TXT_YEAR_VAL = excdatestr.Year.ToString();
                                }
                                catch
                                {
                                    /*TXT_TGL_.Text = "";
                                    DDL_BLN_.SelectedValue = "";
                                    TXT_YEAR_.Text = "";*/

                                    TXT_TGL_VAL = "";
                                    DDL_BLN_ = "";
                                    TXT_YEAR_ = "";
                                }
                                result.Add(TXT_TGL_, TXT_TGL_VAL);
                                result.Add(DDL_BLN_, DDL_BLN_VAL);
                                result.Add(TXT_YEAR_, TXT_YEAR_VAL);
                                //--------------------------------------------------------------------------------------//
                            }
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk cash bank sampe liabilities net worth, lihat excel !!!!!!
                    for (int m = 66; m < 70; m++)
                    {
                        string TXT_ = "";
                        string TXT_VAL = "";
                        /// Start Read Neraca
                        /// 
                        //for (int n=4;n<=conn.GetRowCount();n++)
                        for (int n = 5; n < 36; n++)
                        {
                            string vRange = ((char)m).ToString() + n;
                            Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vRange, vRange);
                            //System.Web.UI.WebControls.TextBox TXT_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_" + vRange);
                            TXT_ = "TXT_" + vRange;
                            //TXT_.Text = formatMoney_ind(excelCell.Value.ToString());
                            try
                            {
                                //TXT_.Text = formatMoney_ind(excelCell.Value2.ToString());
                                //TXT_.Text = GlobalTools.MoneyFormat(excelCell.Value2.ToString());
                                TXT_VAL = formatMoney_ind(excelCell.Value2.ToString());
                            }
                            catch
                            {
                                //TXT_.Text = "0";
                                TXT_VAL = "0";
                            }

                            result.Add(TXT_, TXT_VAL);
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                /*}
                catch (Exception e)
                {
                    string resultmsg = e.Message;
                }
                finally
                {*/
                    excelWorkbook.Close(null, null, null);
                    excelApp.Workbooks.Close();
                    excelApp.Application.Quit();
                    excelApp.Quit();

                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets); 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    //excelSheets = null; 
                    excelWorkbook = null;
                    excelApp = null;

                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }

                //}

                results = result;
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return returnMsg;
        }

        string IWord.Neraca_KMK_KI_MediumASPXviewExcel_LabaRugi(string directori, string regno, string userid, out Dictionary<string, string> results)
        {
            Dictionary<string, string> resultD = new Dictionary<string, string>();

            string vPath;
            string result = "";

            try
            {
                //TODO : Jangan di hardcode !!!
                conn.QueryString = "select xls_dir+''+fu_filename as filexls from CA_FILEUPLOADXL where fu_filename = '" + regno + "-" + userid + "-" + directori + "'";
                conn.ExecuteQuery();
                vPath = conn.GetFieldValue("filexls");


                Microsoft.Office.Interop.Excel.Application excelAppIS = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkbookIS = null;

                /////////////////////////////////
                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();
                /////////////////////////////////

                /////////////////////////////////////////////////////////////////
                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);
                ////////////////////////////////////////////////////////////////

                try
                {
                    // Set the culture and UI culture to the browser's accept language
                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelAppIS = new Microsoft.Office.Interop.Excel.Application();
                    excelAppIS.Visible = false;
                    excelAppIS.DisplayAlerts = false;

                    ////////////////////////////////////////////////////////////////
                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);
                    ////////////////////////////////////////////////////////////////

                    /// Save process into database
                    /// 					
                    //SupportTools.saveProcessExcel(excelAppIS, newId, orgId, conn);
                    vPath = vPath.Replace(".XLSX", ".XLSX");
                    excelWorkbookIS = excelAppIS.Workbooks.Open(vPath,
                        0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        false, false, 0, true);
                    Microsoft.Office.Interop.Excel.Sheets excelSheets = excelWorkbookIS.Worksheets;
                    string currentSheet = "LOS";
                    Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheets.get_Item(currentSheet);
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk date periode, number of months lihat excel !!!!!!!!!!!
                    for (int i = 66; i < 70; i++)
                    {
                        string TXT_LBRG_ = "";
                        string TXT_LBRG_VAL = "";

                        for (int j = 37; j < 41; j++)
                        {
                            string vtmp = ((char)i).ToString() + j; //i=66 diconvert ke ascci jd huruf B, di concat dgn j hasilnya B1,B2,C1,C2
                            Microsoft.Office.Interop.Excel.Range excelLBRG1 = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vtmp, vtmp);
                            //System.Web.UI.WebControls.TextBox TXT_LBRG_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_LBRG_" + vtmp);
                            TXT_LBRG_ = "TXT_LBRG_" + vtmp;
                            if (j == 37)
                            {
                                try
                                {
                                    //TXT_LBRG_.Text = excelLBRG1.Text.ToString();
                                    //TXT_LBRG_.Text = excelLBRG1.Value2.ToString();
                                    TXT_LBRG_VAL = excelLBRG1.Value2.ToString();
                                }
                                catch
                                {
                                    //TXT_LBRG_.Text = "";
                                    TXT_LBRG_VAL = "";
                                }
                            }
                            else
                            {
                                try
                                {
                                    //TXT_LBRG_.Text = excelLBRG1.Value2.ToString();
                                    TXT_LBRG_VAL = excelLBRG1.Value2.ToString();
                                }
                                catch
                                {
                                    //TXT_LBRG_.Text = "";
                                    TXT_LBRG_VAL = "";
                                }
                            }

                            resultD.Add(TXT_LBRG_, TXT_LBRG_VAL);
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                    // loop utk NET SALES sampe   % OF SALES, lihat excel !!!!!!
                    for (int m = 66; m < 70; m++)
                    {
                        string TXT_LBRG_ = "";
                        string TXT_LBRG_VAL = "";

                        //for (int n=4;n<=conn.GetRowCount();n++) 
                        for (int n = 41; n < 62; n++)
                        {
                            int a = n;
                            string vRange = ((char)m).ToString() + n;
                            Microsoft.Office.Interop.Excel.Range excelLBRG2 = (Microsoft.Office.Interop.Excel.Range)excelWorksheet.get_Range(vRange, vRange);
                            //System.Web.UI.WebControls.TextBox TXT_LBRG_ = (System.Web.UI.WebControls.TextBox)thisPage.FindControl("TXT_LBRG_" + vRange);
                            TXT_LBRG_ = "TXT_LBRG_" + vRange;

                            try
                            {
                                //TXT_LBRG_.Text = formatMoney_ind(excelLBRG2.Value2.ToString());
                                TXT_LBRG_VAL = formatMoney_ind(excelLBRG2.Value2.ToString());
                                //TXT_LBRG_.Text = GlobalTools.MoneyFormat(excelLBRG2.Value2.ToString());
                            }
                            catch (Exception e)
                            {
                                //TXT_LBRG_.Text = "0";
                                TXT_LBRG_VAL = "0";
                                result = e.Message;
                            }

                            resultD.Add(TXT_LBRG_, TXT_LBRG_VAL);
                        }
                    }
                    /*--------------------- separator ---------------------------------------------------------------*/
                }
                catch (Exception e)
                {
                    result = e.Message;
                }
                finally
                {
                    try
                    {
                        excelWorkbookIS.Close(null, null, null);
                    }
                    catch (Exception e)
                    {
                        result = e.Message;
                    }

                    try
                    {
                        excelAppIS.Workbooks.Close();
                    }
                    catch (Exception e)
                    {
                        result = e.Message;
                    }

                    try
                    {
                        excelAppIS.Application.Quit();
                    }
                    catch (Exception e)
                    {
                        result = e.Message;
                    }

                    try
                    {
                        excelAppIS.Quit();
                    }
                    catch (Exception e)
                    {
                        result = e.Message;
                    }

                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheets); 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbookIS);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelAppIS);
                    //excelSheets = null; 
                    excelWorkbookIS = null;
                    excelAppIS = null;

                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }

                results = resultD;
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return result;
        }

        string IWord.AppraisalNewASPXReadExcel(string filename, string templateid, string regno, string curef, string clseq)
        {
            string resultmsg = "";

            try
            {

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess) orgId.Add(thisProcess);

                System.Data.DataTable dt1, dt2;

                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                    System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;

                    Process[] newProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in newProcess) newId.Add(thisProcess);

                    //Save process into database
                    //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);
                    filename = filename.Replace(".XLSX", ".XLSX");
                    excelWorkBook = excelApp.Workbooks.Open(filename,
                        0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        false, false, 0, true);

                    excelSheet = excelWorkBook.Worksheets;

                    //Loop for Template Master
                    conn.QueryString = "SELECT SHEET_ID, SHEET_SEQ, STOREDPROCEDURE FROM APPRAISALNEW_TEMPLATE_MASTER WHERE TEMPLATE_ID = '" + templateid + "'";
                    conn.ExecuteQuery();

                    dt1 = conn.GetDataTable().Copy();

                    if (dt1.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            string sheetid = dt1.Rows[i][0].ToString().Trim();
                            string sheetseq = dt1.Rows[i][1].ToString().Trim();
                            string proc = dt1.Rows[i][2].ToString().Trim();

                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheetid);

                            //Loop for Template Detail
                            conn.QueryString = "SELECT CELL_ROW, CELL_COL, DB_FIELD FROM APPRAISALNEW_TEMPLATE_DETAIL WHERE TEMPLATE_ID = '" + templateid +
                                "' AND SHEET_ID = '" + sheetid + "' AND SHEET_SEQ = '" + sheetseq + "' ORDER BY SEQ";
                            conn.ExecuteQuery();
                            dt2 = conn.GetDataTable().Copy();
                            int n = dt2.Rows.Count;
                            object[] par;
                            par = new object[n];
                            object[] dttype;
                            dttype = new object[n];

                            if (dt2.Rows.Count > 0)
                            {
                                for (int j = 0; j < dt2.Rows.Count; j++)
                                {
                                    string xrow = dt2.Rows[j][0].ToString().Trim();
                                    string xcol = dt2.Rows[j][1].ToString().Trim();
                                    string datatype = dt2.Rows[j][2].ToString().Trim(); //data type
                                    string cell_value;
                                    string xcell = xcol + xrow;

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(xcell, xcell);
                                    if (excelCell != null)
                                    {
                                        cell_value = excelCell.Value2.ToString();
                                        par[j] = (string)cell_value;
                                        dttype[j] = (string)datatype;
                                    }
                                }

                                //Construct Query
                                string query = "EXEC " + proc + " '" + regno + "', '" + curef + "', '" + clseq + "', ";
                                for (int k = 0; k < n; k++)
                                {
                                    if (dttype[k].ToString() == "C")
                                        query = query + "'" + par[k].ToString() + "'";
                                    else if (dttype[k].ToString() == "N")
                                        query = query + "" + par[k].ToString() + "";

                                    if (k < n - 1)
                                        query = query + ", ";
                                }

                                //Run Query
                                conn.QueryString = query;
                                conn.ExecuteQuery();

                                //Show Success Message
                                resultmsg = "Upload Sucessful! Insert Result Sucessful!";
                            }
                        }

                        return resultmsg;
                    }
                }
                catch (Exception ex)
                {
                    /*
                    LBL_STATUS.ForeColor = Color.Red;
                    LBL_STATUSREPORT.ForeColor = Color.Red;
                    LBL_STATUS.Text = "Upload Failed!";
                    LBL_STATUSREPORT.Text = ex.Message + "\n" + ex.StackTrace;
                    */
                    //Response.Write("<!--" + ex.Message + "\n" + ex.StackTrace + "-->");
                    resultmsg = ex.Message;
                }
                finally
                {
                    if (excelWorkBook != null)
                    {
                        excelWorkBook.Close(true, filename, null);
                        excelWorkBook = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.Workbooks.Close();
                        excelApp.Application.Quit();
                        excelApp = null;
                    }
                }

                try
                {
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
                catch (Exception a)
                {
                    resultmsg = a.Message;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return resultmsg;
        }

        string IWord.CreditProposalMainExport_Word(string regno, string userid, string var_idExport1, string var_idExport2)
        {

            System.Data.DataTable dt_field = null;

            /*
            string areaid=null;
            string branchid=null;
            */

            string data_id = null;
            string prgid = null;
            string fileNm = string.Empty;
            string fileIn = string.Empty;
            string fileOut = string.Empty;
            object objValue = null;
            object objType = Type.Missing;
            string mStatus = string.Empty;

            ArrayList orgId = new ArrayList();
            ArrayList newId = new ArrayList();

            try
            {

                /// Mengambil application root
                /// 
                conn.QueryString = "select APP_ROOT from APP_PARAMETER";
                conn.ExecuteQuery();
                string vAPP_ROOT = conn.GetFieldValue("APP_ROOT");

                //--- INIT DATA_ID => 
                conn.QueryString = "Select nota_id,programid from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();
                data_id = conn.GetFieldValue("nota_id").ToString();
                prgid = conn.GetFieldValue("programid").ToString();

                /// Mengambil nilai parameter
                /// 
                conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORDPROC '" + data_id + "', '" + prgid + "'";
                conn.ExecuteQuery();
                System.Data.DataTable dtProc = new System.Data.DataTable();
                dtProc = conn.GetDataTable().Copy();

                if (conn.GetRowCount() == 0)
                {
                    //GlobalTools.popMessage(thisPage, "Data Referensi nota analisa word kosong!");
                    mStatus = "Data Referensi nota analisa word kosong!";
                    //return ;
                }

                string nota = data_id;											// nama file hasil export
                string sheet = conn.GetFieldValue("nota_id");
                //string path = vAPP_ROOT + conn.GetFieldValue("nota_PATH");	// directory WORD hasil export			
                string path = conn.GetFieldValue("nota_PATH");
                string file_xls = nota + ".docx";								// nama file WORD template
                string template = conn.GetFieldValue("nota_id");				// directory WORD template
                string template_path = conn.GetFieldValue("TEMPLATE_PATH");				// directory WORD template

                string url = conn.GetFieldValue("nota_URL");					// url (link) untuk download

                string[] procedure_name = new string[100];

                /*
                for(int den=0; den < dtProc.Rows.Count; den++) 
                {
                    procedure_name[den] = conn.GetFieldValue(den, "STOREDPROCEDURE");
                }
                */


                fileNm = regno + "-" + nota + "-" + userid + ".DOCX";

                object objFileIn = template_path + file_xls;
                object objFileOut = path + fileNm;

                //fileResult = url + fileNm;



                /// Cek apakah file templatenya (input) ada atau tidak
                /// 
                //if (!File.Exists(template + file_xls)) 
                if (!File.Exists(template_path + file_xls))
                {
                    //GlobalTools.popMessage(thisPage, "File Template tidak ada!");
                    mStatus = "File Template tidak ada!";
                    //return ;
                }

                /// Cek direktori untuk menyimpan file hasil export (output)
                /// 
                if (!Directory.Exists(path))
                {
                    // create directory if does not exist
                    Directory.CreateDirectory(path);
                }


                /// dapatkan semua fields to populate
                /// 			


                object oMissingObject = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document wordDoc = null;


                Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);


                // Always already when using Export Excel file format					
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");


                wordApp = new Application();
                wordApp.Visible = false;

                //Collecting Existing Winword in Taskbar 

                Process[] newProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in newProcess)
                    newId.Add(thisProcess);

                /// Save process into database
                /// 					
                //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);	


                wordDoc = wordApp.Documents.Open(ref objFileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                    ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);
                wordDoc.Activate();
                Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;

                int procedure_cnt = 0;  // for the moment, we handle 9 store procedures			
                string exe_procedure;

                object oCell;
                string tempField;
                object sField;
                string strObject;


                /*
                conn.QueryString = " exec CP_EXPORT_NOTA_ANALISA_WORDPROC '" + data_id + "','" + prgid + "'"; 
                conn.ExecuteQuery();
			
			
			
                for(int den=0; den < conn.GetRowCount(); den++) 
                {
                    procedure_name[den] = conn.GetFieldValue(den, "STOREDPROCEDURE");
                }
                */


                for (procedure_cnt = 0; procedure_cnt < dtProc.Rows.Count; procedure_cnt++)
                {

                    exe_procedure = dtProc.Rows[procedure_cnt]["STOREDPROCEDURE"].ToString();
                    if (Strings.Len(exe_procedure.Trim()) == 0) continue;

                    /*
                    conn.QueryString = "Select SEQ, nota_COL, nota_ROW, nota_FIELD, [group] from nota_analisa_detail " + 
                        " where nota_ID = '" + nota + 
                        "' and category = " + procedure_cnt.ToString() +
                        " order by SEQ";
                    conn.ExecuteQuery();
                    */
                    /*
                    conn.QueryString = "select d.NOTA_ID,d.SEQ,d.NOTA_COL,d.NOTA_ROW," +
                                    "d.NOTA_FIELD,d.[DESCRIPTION],d.[Group],d.category," +
                                    "p.STOREDPROCEDURE " +
                                    " from  nota_analisa_detail d left join rfnotaanalisaproc p on " +
                                    "d.nota_id = p.nota_id and " +
                                    "d.category = p.seq " +
                                    "where d.nota_id = '" + nota + "' and category = " + 
                                     procedure_cnt.ToString()        
                                     + " order by d.SEQ ";
                                     */
                    conn.QueryString = "select d.NOTA_ID,d.SEQ,d.NOTA_COL,d.NOTA_ROW," +
                        "d.NOTA_FIELD,d.[DESCRIPTION],d.[Group],d.category," +
                        "p.STOREDPROCEDURE " +
                        " from  nota_analisa_detail d left join rfnotaanalisaproc p on " +
                        "d.nota_id = p.nota_id and " +
                        "d.category = p.seq " +
                        "where d.nota_id = '" + nota + "' and p.STOREDPROCEDURE = '" +
                        exe_procedure.ToString() + "' order by d.SEQ ";
                    conn.ExecuteQuery();
                    dt_field = conn.GetDataTable().Copy();


                    conn.QueryString = " exec " + exe_procedure + " '" + regno + "'";
                    conn.ExecuteQuery();

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {

                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {

                            try
                            {
                                oCell = dt_field.Rows[i]["nota_col"];
                                tempField = dt_field.Rows[i]["nota_field"].ToString();
                                sField = dt_field.Rows[i]["nota_field"].ToString();


                                objValue = conn.GetFieldValue(j, tempField);

                                //if(wordBookMark.Exists(oCell.ToString())) 
                                if (wordBookMark.Exists(sField.ToString()))
                                {

                                    if (dt_field.Rows[i]["Group"].ToString() != "0") strObject = objValue.ToString();
                                    else strObject = objValue.ToString() + "\n";

                                    //Word.Bookmark oBook = wordBookMark.Item(ref oCell);
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref sField);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;


                                }
                            }
                            catch { }

                        }  // end of for i loop				

                    }  // end of j 
                    // close the objects					
                }  // end of procedure_cnt


                try
                {
                    /// Save file fisik hasil export
                    /// 
                    //excelWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

                    wordDoc.SaveAs(ref objFileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                        ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));

                    conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                    conn.ExecuteQuery();
                    mStatus = "Export Succesfully";

                    /*  mungkin return value didalam calling fucntion ....
                        /// Save data file hasil export ke database
                        /// 
                        Conn.QueryString = "exec RPT_EXPORT_DATAANALYSIS '" + 
                            data_id + "', '" + 
                            fileNm + "', '" + 
                            var_userid + " ', '1', 'R-C'";
                        Conn.ExecuteNonQuery();					
                        */


                }
                catch  // (Exception exp2)
                {
                    // LBL_STATUS_EXPORT.Text = "Export File gagal!";
                    //LBL_STATUSEXPORT.Text = exp2.ToString();
                    //return;
                }

                // try to close word dulu ...
                try
                {
                    if (wordDoc != null)
                    {
                        ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordDoc = null;
                    }
                    if (wordApp != null)
                    {
                        ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordApp = null;
                    }
                }
                catch { }


                /// Kill process
                /// 
                try
                {

                    // Killing Proses after Export
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }

                    } // end x		
                }
                catch { }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateNotaWord(string regno, string userid, string sessionfullName, string branchName, string ddl_manualSelectedValue, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue)
        {
            string szUser = ddl_manualSelectedValue;
            string fileNm = string.Empty;
            string fileIn = string.Empty;
            string fileOut = string.Empty;
            string mStatus = string.Empty;
            string mNotaNumber = string.Empty;
            short Step = 0;
            object objType = Type.Missing;
            bool bSukses = true;
            object objValue = null;
            string var_user = userid;
            string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
            string var_idExport2 = DDL_KETENTUANSelectedValue;

            try
            {

                System.Data.DataTable dt_field = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_doc = nota + ".DOCX";
                    string url = conn.GetFieldValue("NOTA_URL");
                    string b_unit = conn.GetFieldValue("B_UNIT");
                    string drill = conn.GetFieldValue("DRILL");

                    int iItem = 0;

                    object oMissingObject = System.Reflection.Missing.Value;

                    Microsoft.Office.Interop.Word.Application wordApp = null;
                    Microsoft.Office.Interop.Word.Document wordDoc = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    //Collecting Existing Winword in Taskbar

                    Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {

                        wordApp = new Application();
                        wordApp.Visible = false;

                        //Collecting Existing Winword in Taskbar 

                        Process[] newProcess = Process.GetProcessesByName("WINWORD");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 					
                        //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);

                        iItem = 0;

                        fileNm = regno + "-" + nota + "-" + var_user + ".DOCX";

                        object objFileIn = path + file_doc;
                        object objFileOut = path + fileNm;

                        //fileResult = url + fileNm;

                        wordDoc = wordApp.Documents.Open(ref objFileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                            ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                        wordDoc.Activate();
                        Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;

                        #region Step Fill General Info
                        // Step 1 - mask out the old codes .... dangerous 
                        //conn.QueryString = "Select * from NOTA_ANALISA_DETAIL where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD '" + branchName + "', '" + sessionfullName + "', '" + regno + "', '" + var_idExport2 + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][2];
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    try
                                    {
                                        strObject = Convert.ToDateTime(objValue).ToShortDateString();
                                    }
                                    catch (Exception e)
                                    {
                                        mStatus = e.Message;
                                    }
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Credit Rating Summary

                        //conn.QueryString = "Select * from NOTA_ANALISA_DETAIL2 where NOTA_ROW = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL2 where NOTA_ROW = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD1_2A '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][2];
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                        }

                        //conn.QueryString = "Select * from NOTA_ANALISA_DETAIL2 where NOTA_ROW = 0 and NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL2 where NOTA_ROW = 0 and NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD1_2B '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][2];
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Yang Memutuskan
                        conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL9 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD9 '" + regno + "', '" + var_user + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][2];
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Ratio
                        conn.QueryString = "Select NOTA_ID,SEQ,STEP,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL3 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        if (drill == "0")
                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD3_1 '" + regno + "', '0'";
                        else
                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD3_2 '" + regno + "'";

                        conn.ExecuteQuery();

                        Step = 0;
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                Step = (short)dt_field.Rows[i][2];

                                if (Step == j + 1)
                                {
                                    object Cell = dt_field.Rows[i][3];
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    string strObject = objValue.ToString();

                                    if (wordBookMark.Exists(Cell.ToString()))
                                    {
                                        Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                        oBook.Select();
                                        oBook.Range.Text = strObject;
                                    }
                                    iItem++;
                                }
                            }
                        }

                        // fill projection ratio (if exist)
                        if (drill == "0")
                        {
                            conn.QueryString = "Select NOTA_ID,SEQ,STEP,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL3 where SEQ = 3 and NOTA_ID = '" + nota + "' order by NOTA_ID";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD3_1 '" + regno + "', '1'";
                            conn.ExecuteQuery();

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    object Cell = dt_field.Rows[i][3];
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    string strObject = objValue.ToString();

                                    if (wordBookMark.Exists(Cell.ToString()))
                                    {
                                        Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                        oBook.Select();
                                        oBook.Range.Text = strObject;
                                    }
                                    iItem++;
                                }
                            }
                        }
                        #endregion
                        #region Step Fill Aspek
                        conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL2 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA2 '" + regno + "'";
                        conn.ExecuteQuery();

                        string szField = string.Empty;
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            szField = conn.GetFieldValue(j, "Control");

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][2];
                                string Field = dt_field.Rows[i][4].ToString();

                                //objValue = conn.GetFieldValue(Field);

                                if (Field == szField)
                                {
                                    objValue = conn.GetFieldValue(j, "Nilai");

                                    if (objValue.ToString() == "1")
                                        objValue = "X";
                                    else if (objValue.ToString() == "0")
                                        objValue = string.Empty;

                                    string strObject = objValue.ToString();

                                    if (wordBookMark.Exists(Cell.ToString()))
                                    {
                                        Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                        oBook.Select();
                                        oBook.Range.Text = strObject;
                                    }

                                    iItem++;

                                    break;
                                }
                            }
                        }
                        #endregion
                        #region Step Fill BU Signature
                        // Step 1
                        conn.QueryString = "Select NOTA_ID,SEQ,STEP,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL10 where Step = 1 and NOTA_ID = '" + nota + "'";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_BU '" + szUser + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][3];
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill RM Signature
                        // Step 1
                        conn.QueryString = "Select NOTA_ID,SEQ,STEP,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL10 where Step = 2 and NOTA_ID = '" + nota + "'";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_RM '" + szUser + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][3];
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill BU FRONT Signature
                        // Step 1
                        conn.QueryString = "Select NOTA_ID,SEQ,STEP,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL10 where Step = 3 and NOTA_ID = '" + nota + "'";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_BUFRONT '" + szUser + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][3];
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill RM FRONT Signature
                        // Step 1
                        conn.QueryString = "Select NOTA_ID,SEQ,STEP,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION] from NOTA_ANALISA_DETAIL10 where Step = 4 and NOTA_ID = '" + nota + "'";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_RMFRONT '" + szUser + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                object Cell = dt_field.Rows[i][3];
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                string strObject = objValue.ToString();

                                if (wordBookMark.Exists(Cell.ToString()))
                                {
                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref Cell);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }

                                iItem++;
                            }
                        }
                        #endregion

                        if (iItem > 0)
                        {
                            wordDoc.SaveAs(ref objFileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                                ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {
                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();
                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {

                        if (wordDoc != null)
                        {
                            ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            wordDoc = null;
                        }
                        if (wordApp != null)
                        {
                            ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                            wordApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        mStatus = e.Message;
                    }
                    // Killing Proses after Export					
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateNotaExcel(string regno, string userid, string SessionBranchName, string SessionFullName, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
            string var_idExport2 = DDL_KETENTUANSelectedValue;
            string szUser = ddl_manualSelectedValue;
            string fileNm = string.Empty;
            string fileIn = string.Empty;
            string fileOut = string.Empty;
            string mNotaNumber = string.Empty;
            short Step = 0;
            object objPaste = null;
            object objCopy = null;
            bool bSukses = true;
            object objValue = null;
            object objType = Type.Missing;
            string mStatus = string.Empty;
            int iItem = 0;
            int iItemOther = 0;
            int iItemPosition = 0;
            int m_Row = 0;

            try
            {
                System.Data.DataTable dt_field = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    //Collecting Existing Excel in Taskbar

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        //Collectiong Existing Excel in Taskbar

                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 					
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);


                        fileIn = path + file_xls;

                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;

                        Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                        var_idExport2 = string.Empty;
                        fileNm = regno + "-" + nota + "-" + userid + ".XLSX";
                        fileOut = path + fileNm;

                        // Sheet ANALISA
                        #region Step Fill General Info
                        // Step 1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA '" + SessionBranchName + "', '" + SessionFullName + "', '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]);
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Multi Ketentuan Kredit
                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_0 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_0 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][4]);

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(),
                                    "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][3].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][5].ToString();

                                if (Field == "NUM")
                                    objValue = j + 1;
                                else
                                    objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Hubungan dengan Bank Mandiri 1
                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_1 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_1 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();
                                if (Field == "NUM")
                                    objValue = j + 1;
                                else
                                    objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Hubungan dengan Bank Mandiri 2
                        // Step 2.2
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_2 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_2 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();
                                if (Field == "NUM")
                                    objValue = j + 1;
                                else
                                    objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;
                            }
                        }
                        #endregion
                        #region Step Fill Hubungan dengan Bank Lain
                        // Step 2.3
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_3 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_3 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;
                            }
                        }
                        #endregion
                        #region Step Fill Aspek
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL2 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA2 '" + regno + "'";
                        conn.ExecuteQuery();

                        string szField = string.Empty;
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            //szField = conn.GetFieldValue(j ,"Control");

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                try
                                {
                                    string Col = dt_field.Rows[i][2].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][4].ToString();

                                    objValue = conn.GetFieldValue(j, Field);
                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;
                                }
                                catch { }
                                /*
                                if (Field==szField)
                                {
                                    objValue = conn.GetFieldValue(j, "Nilai");

                                    if(objValue.ToString() == "1") 
                                        objValue = "X";
                                    else if(objValue.ToString() == "0")
                                        objValue = string.Empty;


                                    Excel.Range excelCell = (Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;

                                    break;
                                }*/
                            }
                        }
                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA2E '" + regno + "'";
                        conn.ExecuteQuery();

                        szField = string.Empty;
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            szField = conn.GetFieldValue(j, "Control");

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();
                                if (Field == szField)
                                {
                                    objValue = conn.GetFieldValue(j, "Nilai");

                                    if (objValue.ToString() == "1")
                                        objValue = "X";
                                    else if (objValue.ToString() == "0")
                                        objValue = string.Empty;

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;

                                    break;
                                }
                            }
                        }

                        #endregion
                        #region Step Fill Ratio
                        conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL3 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA3 '" + regno + "'";
                        conn.ExecuteQuery();

                        Step = 0;
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                Step = (short)dt_field.Rows[i][2];

                                if (Step == j + 1)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                        }
                        #endregion
                        #region Step Fill Signature
                        // Step 2.2
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL9 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA9 '" + regno + "', '" + userid + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;
                            }
                        }
                        #endregion
                        #region Step Fill BU Signature
                        // Step 1
                        conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL10 where Step = 1 and NOTA_ID = '" + nota + "'";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_BU '" + szUser + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][3].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill RM Signature
                        // Step 1
                        conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL10 where Step = 2 and NOTA_ID = '" + nota + "'";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_RM '" + szUser + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][3].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion

                        if (iItem > 0)
                        {
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {
                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();
                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }
                    try
                    {
                        // Killing Proses after Export
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateNew(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;

                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                bool bSukses = true;
                object objValue = null;
                object objType = Type.Missing;
                int iItem = 0;
                int m_Row = 0;


                System.Data.DataTable dt_field = null;
                System.Data.DataTable dt_field2 = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    if (var_idExport2 != string.Empty)
                    {
                        try
                        {
                            excelApp = new Microsoft.Office.Interop.Excel.Application();
                            excelApp.Visible = false;
                            excelApp.DisplayAlerts = false;

                            Process[] newProcess = Process.GetProcessesByName("EXCEL");
                            foreach (Process thisProcess in newProcess)
                                newId.Add(thisProcess);

                            /// Save process into database
                            /// 
                            //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                            fileIn = path + file_xls;

                            excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                                false, false, 0, true);

                            excelSheet = excelWorkBook.Worksheets;

                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);
                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheetWork = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item("WORK");

                            fileNm = regno + "-" + nota + "-" + var_Name + "-" + prog_code + "-" + var_user + ".XLSX";
                            fileOut = path + fileNm;
                            //fileResult = url + fileNm;

                            // Sheet NEW
                            #region Step Fill MultiKetentuan New

                            // Step 2.1

                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL7 where step = 0 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_0 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            m_Row = 0;

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]);
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }


                            // Step 2.2

                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL7 where step = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec DE_TOTALEXPOSURE '" + regno + "'";
                            conn.ExecuteQuery(300);

                            m_Row = 0;

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]);
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }

                            #region Include Agunan
                            // Step 2.1
                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL4 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field2 = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA4_1 '" + regno + "','" + var_idExport2 + "','01'";
                            conn.ExecuteQuery();
                            #region OLD_INSERTION
                            /***** MASK OUT OLEH DENNY ------
						if((conn.GetRowCount()>1) && (dt_field2.Rows.Count>1))
						{
							#region Inserting
							// Ini insert untuk Sheet New
							iItemPosition = Convert.ToInt32(dt_field2.Rows[0][4]) + 1;

							//Prepare Empty Row with all Formats
							for(int k = 0; k < conn.GetRowCount()-1; k++)
							{
								Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
								excelRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

								Excel.Range excelRangeWork = excelWorkSheetWork.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
								excelRangeWork.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
								
								iItemOther = iItemPosition + k;
							}

							Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition - 1).ToString().Trim(), "IV" + (iItemPosition - 1).ToString().Trim());
							objCopy = exlCopy.Copy(objType);

							Excel.Range exlPaste = excelWorkSheet.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
							objPaste = exlPaste.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

							Excel.Range exlCopyWork = excelWorkSheetWork.get_Range("A" + (iItemPosition - 1).ToString().Trim(), "IV" + (iItemPosition - 1).ToString().Trim());
							objCopy = exlCopyWork.Copy(objType);

							Excel.Range exlPasteWork = excelWorkSheetWork.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
							objPaste = exlPasteWork.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
							#endregion
						}
						****/
                            #endregion
                            for (int k = 0; k < conn.GetRowCount(); k++)
                            {
                                if (k > 0) m_Row = m_Row + 1;

                                for (int l = 0; l < dt_field2.Rows.Count; l++)
                                {
                                    string Col1 = dt_field2.Rows[l][3].ToString().Trim();
                                    int Row1 = Convert.ToInt32(dt_field2.Rows[l][4]) + m_Row;
                                    string Cell1 = Col1 + Row1.ToString().Trim();
                                    string Field1 = dt_field2.Rows[l][5].ToString();
                                    if (Field1 == "NUM")
                                        objValue = k + 1;
                                    else
                                        objValue = conn.GetFieldValue(k, Field1);

                                    Microsoft.Office.Interop.Excel.Range excelCell1 = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell1, Cell1);
                                    excelCell1.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion
                            #endregion

                            if (iItem > 0)
                            {
                                excelWorkSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                                excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                     Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                                bSukses = true;
                            }
                            else
                                bSukses = false;

                            if (bSukses)
                            {
                                // Maintenance Table Nota_Export

                                if (var_idExport2 == string.Empty)
                                {
                                    conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                                }
                                else
                                {
                                    conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";
                                }
                                conn.ExecuteQuery();

                                mStatus = "Export Succesfully";

                            }
                            else
                            {
                                mStatus = "No Data to Export";
                            }
                        }
                        catch (Exception e)
                        {
                            //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                            mStatus = e.Message;
                        }
                        finally
                        {
                            if (excelWorkBook != null)
                            {
                                excelWorkBook.Close(true, fileOut, null);
                                excelWorkBook = null;
                            }
                            if (excelApp != null)
                            {
                                excelApp.Workbooks.Close();
                                excelApp.Application.Quit();
                                excelApp = null;
                            }
                        }
                        try
                        {
                            for (int x = 0; x < newId.Count; x++)
                            {
                                Process xnewId = (Process)newId[x];

                                bool bSameId = false;
                                for (int z = 0; z < orgId.Count; z++)
                                {
                                    Process xoldId = (Process)orgId[z];

                                    if (xnewId.Id == xoldId.Id)
                                    {
                                        bSameId = true;
                                        break;
                                    }
                                }
                                if (bSameId)
                                {
                                    try
                                    {
                                        xnewId.Kill();
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            mStatus = e.Message;
                        }
                    }
                    else
                    {
                        //GlobalTools.popMessage(thisPage, "Ketentuan tidak boleh kosong!");
                        mStatus = "Ketentuan tidak boleh kosong!";
                    }
                    //GlobalTools.popMessage(this, "Ketentuan must be filled");
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateExist(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;

                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                bool bSukses = true;
                object objValue = null;
                object objType = Type.Missing;
                int iItem = 0;
                int iItemOther = 0;
                int iItemPosition = 0;
                int m_Row = 0;

                System.Data.DataTable dt_field = null;
                System.Data.DataTable dt_field1 = null;
                System.Data.DataTable dt_field2 = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    if (var_idExport2 != string.Empty)
                    {
                        try
                        {
                            excelApp = new Microsoft.Office.Interop.Excel.Application();
                            excelApp.Visible = false;
                            excelApp.DisplayAlerts = false;

                            Process[] newProcess = Process.GetProcessesByName("EXCEL");
                            foreach (Process thisProcess in newProcess)
                                newId.Add(thisProcess);

                            /// Save process into database
                            /// 
                            //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                            fileIn = path + file_xls;

                            excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                                false, false, 0, true);

                            excelSheet = excelWorkBook.Worksheets;

                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);
                            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheetWork = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item("WORK");

                            fileNm = regno + "-" + nota + "-" + var_Name + "-" + prog_code + "-" + var_user + ".XLSX";
                            fileOut = path + fileNm;
                            //fileResult = url + fileNm;

                            #region Step Fill Header
                            // Step 2.1

                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL6 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_1 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]);
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion
                            #region Step Fill Withdrawal
                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL7 where Step = 0 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_2 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            dt_field1 = conn.GetDataTable().Copy();

                            m_Row = 0;

                            for (int j = 0; j < dt_field1.Rows.Count; j++)
                            {
                                if (j > 0) m_Row = m_Row + 26;

                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = dt_field1.Rows[j][Field];

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion
                            #region Step Fill Renewal
                            // Step 2.1

                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL7 where Step = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_3 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion
                            #region Step Fill Limit Changes
                            // Step 2.1

                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL7 where Step = 2 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_4 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion
                            #region Step Fill Change Collateral

                            #region Include Exiting Agunan
                            // Step untuk Exisiting adalah 1 karena belum dapat result query maka step diganti 4
                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL4 where STEP = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field2 = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA4_0 '" + regno + "', '" + var_idExport2 + "','07'";
                            conn.ExecuteQuery();

                            #region INSERT_OLD_ROWS
                            /**** remove by Denny
						if((conn.GetRowCount()>1) && (dt_field2.Rows.Count>1))
						{
							iItemPosition = Convert.ToInt32(dt_field2.Rows[0][4]) + m_Row + 1;

							//Prepare Empty Row with all Formats
							for(int k = 0; k < conn.GetRowCount()-1; k++)
							{
								Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
								excelRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

								Excel.Range excelRangeWork = excelWorkSheetWork.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
								excelRangeWork.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
								
								iItemOther = iItemPosition + k;
							}

							Excel.Range exlCopyOld = excelWorkSheet.get_Range("A" + (iItemPosition - 1).ToString().Trim(), "IV" + (iItemPosition - 1).ToString().Trim());
							objCopy = exlCopyOld.Copy(objType);

							Excel.Range exlPasteOld = excelWorkSheet.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
							objPaste = exlPasteOld.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

							Excel.Range exlCopyWork1 = excelWorkSheetWork.get_Range("A" + (iItemPosition - 1).ToString().Trim(), "IV" + (iItemPosition - 1).ToString().Trim());
							objCopy = exlCopyWork1.Copy(objType);

							Excel.Range exlPasteWork1 = excelWorkSheetWork.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
							objPaste = exlPasteWork1.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
						}	
						***/
                            #endregion
                            //						if((conn.GetRowCount()>1) && (dt_field2.Rows.Count>1))
                            //						{
                            //							iItemPosition = Convert.ToInt32(dt_field2.Rows[0][4]) + m_Row;
                            //
                            //							//Prepare Empty Row with all Formats
                            //							for(int k = 0; k < conn.GetRowCount()-1; k++)
                            //							{
                            //								Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                            //								excelRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //								iItemOther = iItemPosition + k;
                            //							}
                            //
                            //							Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            //							objCopy = exlCopy.Copy(objType);
                            //
                            //							Excel.Range exlPaste = excelWorkSheet.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            //							objPaste = exlPaste.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            //						}
                            for (int k = 0; k < conn.GetRowCount(); k++)
                            {
                                if (k > 0) m_Row = m_Row + 1;

                                for (int l = 0; l < dt_field2.Rows.Count; l++)
                                {
                                    string Col1 = dt_field2.Rows[l][3].ToString().Trim();
                                    int Row1 = Convert.ToInt32(dt_field2.Rows[l][4]) + m_Row;
                                    string Cell1 = Col1 + Row1.ToString().Trim();
                                    string Field1 = dt_field2.Rows[l][5].ToString();
                                    if (Field1 == "NUM")
                                        objValue = k + 1;
                                    else
                                        objValue = conn.GetFieldValue(k, Field1);

                                    Microsoft.Office.Interop.Excel.Range excelCell1 = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell1, Cell1);
                                    excelCell1.Value2 = objValue;

                                    iItem++;
                                }
                            }

                            #endregion
                            #region Include New Agunan
                            // Step 2.1
                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL4 where STEP = 2 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field2 = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA4_1 '" + regno + "', '" + var_idExport2 + "','02'";
                            conn.ExecuteQuery();

                            // init here because we are no longer adding rows in above ....
                            m_Row = 0;

                            iItemPosition = Convert.ToInt32(dt_field2.Rows[0][4]) + m_Row + 1;
                            iItemOther = iItemPosition;

                            //Prepare Empty Row with all Formats
                            #region insert_posisi
                            /***
						for(int k = 0; k < conn.GetRowCount()-1; k++)
						{
							Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
							excelRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

							Excel.Range excelRangeWork = excelWorkSheetWork.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
							excelRangeWork.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
								
							iItemOther = iItemPosition + k;
						}

						Excel.Range exlCopyNew = excelWorkSheet.get_Range("A" + (iItemPosition - 1).ToString().Trim(), "IV" + (iItemPosition - 1).ToString().Trim());
						objCopy = exlCopyNew.Copy(objType);

						Excel.Range exlPasteNew = excelWorkSheet.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
						objPaste = exlPasteNew.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

						Excel.Range exlCopyWork2 = excelWorkSheetWork.get_Range("A" + (iItemPosition - 1).ToString().Trim(), "IV" + (iItemPosition - 1).ToString().Trim());
						objCopy = exlCopyWork2.Copy(objType);

						Excel.Range exlPasteWork2 = excelWorkSheetWork.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
						objPaste = exlPasteWork2.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
						***/
                            #endregion

                            //						if((conn.GetRowCount()>1) && (dt_field2.Rows.Count>1))
                            //						{
                            //							iItemPosition = Convert.ToInt32(dt_field2.Rows[0][4]) + m_Row;
                            //
                            //							//Prepare Empty Row with all Formats
                            //							for(int k = 0; k < conn.GetRowCount()-1; k++)
                            //							{
                            //								Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                            //								excelRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            //								iItemOther = iItemPosition + k;
                            //							}
                            //
                            //							Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            //							objCopy = exlCopy.Copy(objType);
                            //
                            //							Excel.Range exlPaste = excelWorkSheet.get_Range("A"+ iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            //							objPaste = exlPaste.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            //						}
                            for (int k = 0; k < conn.GetRowCount(); k++)
                            {
                                if (k > 0) m_Row = m_Row + 1;

                                for (int l = 0; l < dt_field2.Rows.Count; l++)
                                {
                                    string Col1 = dt_field2.Rows[l][3].ToString().Trim();
                                    int Row1 = Convert.ToInt32(dt_field2.Rows[l][4]) + m_Row;
                                    string Cell1 = Col1 + Row1.ToString().Trim();
                                    string Field1 = dt_field2.Rows[l][5].ToString();
                                    if (Field1 == "NUM")
                                        objValue = k + 1;
                                    else
                                        objValue = conn.GetFieldValue(k, Field1);

                                    Microsoft.Office.Interop.Excel.Range excelCell1 = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell1, Cell1);
                                    excelCell1.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion
                            #region Step Fill Syarat

                            // Step 2.1
                            conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL7 where Step = 4 and NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_7 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][3].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][5].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion

                            #endregion
                            #region Step Fill Perubahan Syarat

                            // Step 2.1
                            conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                            conn.ExecuteQuery();

                            dt_field = conn.GetDataTable().Copy();

                            conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA7_6 '" + regno + "', '" + var_idExport2 + "'";
                            conn.ExecuteQuery();

                            for (int j = 0; j < conn.GetRowCount(); j++)
                            {
                                for (int i = 0; i < dt_field.Rows.Count; i++)
                                {
                                    string Col = dt_field.Rows[i][2].ToString().Trim();
                                    int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                    string Cell = Col + Row.ToString().Trim();
                                    string Field = dt_field.Rows[i][4].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                    excelCell.Value2 = objValue;

                                    iItem++;
                                }
                            }
                            #endregion

                            if (iItem > 0)
                            {
                                excelWorkSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
                                excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                                bSukses = true;
                            }
                            else
                                bSukses = false;

                            if (bSukses)
                            {
                                // Maintenance Table Nota_Export

                                if (var_idExport2 == string.Empty)
                                {
                                    conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                                }
                                else
                                {
                                    conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";
                                }

                                conn.ExecuteQuery();
                                mStatus = "Export Succesfully";
                            }
                            else
                            {
                                mStatus = "No Data to Export";
                            }
                        }
                        catch (Exception e)
                        {
                            //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                            mStatus = e.Message;
                        }
                        finally
                        {
                            if (excelWorkBook != null)
                            {
                                excelWorkBook.Close(true, fileOut, null);
                                excelWorkBook = null;
                            }
                            if (excelApp != null)
                            {
                                excelApp.Workbooks.Close();
                                excelApp.Application.Quit();
                                excelApp = null;
                            }
                        }
                        try
                        {
                            for (int x = 0; x < newId.Count; x++)
                            {
                                Process xnewId = (Process)newId[x];

                                bool bSameId = false;
                                for (int z = 0; z < orgId.Count; z++)
                                {
                                    Process xoldId = (Process)orgId[z];

                                    if (xnewId.Id == xoldId.Id)
                                    {
                                        bSameId = true;
                                        break;
                                    }
                                }
                                if (bSameId)
                                {
                                    try
                                    {
                                        xnewId.Kill();
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        //GlobalTools.popMessage(thisPage, "Ketentuan tidak boleh kosong!");
                        mStatus = "Ketentuan tidak boleh kosong!";
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateSyarat(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;
                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objPaste = null;
                object objCopy = null;
                bool bSukses = true;
                object objValue = null;
                object objType = Type.Missing;
                int iItem = 0;
                int iItemOther = 0;
                int iItemPosition = 0;
                int m_Row = 0;

                System.Data.DataTable dt_field = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = path + file_xls;

                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;

                        Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                        fileNm = regno + "-" + nota + "-" + var_user + ".XLSX";
                        fileOut = path + fileNm;
                        //fileResult = url + fileNm;

                        m_Row = 0;

                        // Sheet SYARAT
                        #region Step Fill Syarat Perjanjian Kredit

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_1 '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion
                        #region Step Fill Syarat Penarikan Kredit

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 2 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_2 '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                iItemOther = iItemPosition + j;

                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion
                        #region Step Fill Syarat Lain

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 3 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_3 '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                iItemOther = iItemPosition + j;

                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion

                        if (iItem > 0)
                        {
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {
                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();
                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateSyarat2(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;
                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objPaste = null;
                object objCopy = null;
                bool bSukses = true;
                object objValue = null;
                object objType = Type.Missing;
                int iItem = 0;
                int iItemOther = 0;
                int iItemPosition = 0;
                int m_Row = 0;

                System.Data.DataTable dt_field = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = path + file_xls;

                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;

                        Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                        fileNm = regno + "-" + nota + "-" + var_user + ".XLSX";
                        fileOut = path + fileNm;
                        //fileResult = url + fileNm;

                        m_Row = 0;

                        // Sheet SYARAT
                        #region Step Fill Syarat Perjanjian Kredit

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 1 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_1 '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion
                        #region Step Fill Syarat Penarikan Kredit
                        #region Step Fill Syarat CL

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 2 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_2A '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                iItemOther = iItemPosition + j;

                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion
                        #region Step Fill Syarat NCL

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 3 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_2B '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                iItemOther = iItemPosition + j;

                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion

                        #endregion
                        #region Step Fill Syarat Lain

                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL8 where SEQ = 4 and NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA8_3 '" + regno + "'";
                        conn.ExecuteQuery();


                        if (conn.GetRowCount() > 1)
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                iItemOther = iItemPosition + j;

                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + (iItemPosition + 1).ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                                Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemPosition + 2).ToString().Trim(), "IV" + (iItemPosition + 3).ToString().Trim());
                                objCopy = exlCopy.Copy(objType);

                                Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemPosition.ToString().Trim());
                                objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                            }
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 2;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion

                        if (iItem > 0)
                        {
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {
                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();
                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (!bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateRata(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;
                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");
                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                bool bSukses = true;
                object objValue = null;


                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    fileNm = regno + "-" + nota + "-" + var_user + ".XLSX";
                    fileOut = path + fileNm;

                    System.Data.DataTable dt_field = null;

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = path + file_xls;

                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;
                        Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                        #region Step Fill Rata-Rata

                        int iItem = 0;
                        // Step 1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL9 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_BANK '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]);
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }

                        #endregion
                        if (iItem > 0)
                        {
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {

                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();
                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }

                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }

                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }

                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateBank(string regno, string userid, string SessionFullName, string SessionBranchName, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;
                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                int m_Row = 0;
                object objPaste = null;
                object objCopy = null;
                object objType = Type.Missing;
                bool bSukses = true;
                object objValue = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    fileNm = regno + "-" + nota + "-" + var_user + ".XLSX";
                    fileOut = path + fileNm;

                    System.Data.DataTable dt_field = null;

                    int iItem = 0;
                    int iItemOther = 0;
                    int iItemPosition = 0;

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = path + file_xls;

                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;
                        Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);


                        #region Step Fill General Info
                        // Step 1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA '" + SessionBranchName + "', '" + SessionFullName + "', '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]);
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Hubungan dengan Bank Mandiri 1
                        // Step 2.1
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_1 where NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_1 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();
                                if (Field == "NUM")
                                    objValue = j + 1;
                                else
                                    objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion
                        #region Step Fill Hubungan dengan Bank Mandiri 2
                        // Step 2.2
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_2 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_2 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();
                                if (Field == "NUM")
                                    objValue = j + 1;
                                else
                                    objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;
                            }
                        }
                        #endregion
                        #region Step Fill Hubungan dengan Bank Lain
                        // Step 2.3
                        conn.QueryString = "Select NOTA_ID, SEQ, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_3 where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA1_3 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][3]) + m_Row;

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }
                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][2].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][3]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;
                            }
                        }
                        #endregion

                        if (iItem > 0)
                        {
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {

                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();
                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }
                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        mStatus = e.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CreditProposalMainp_CreateUrus(string regno, string userid, string DDL_FORMAT_TYPESelectedValue, string DDL_KETENTUANSelectedValue, string ddl_manualSelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                string szUser = ddl_manualSelectedValue;
                string var_user = userid;
                string var_idExport1 = DDL_FORMAT_TYPESelectedValue;
                string var_idExport2 = DDL_KETENTUANSelectedValue;

                conn.QueryString = "select KET_CODE, KET_DESC from KETENTUAN_KREDIT where AP_REGNO = '" + regno + "' and KET_CODE = '" + var_idExport2 + "'";
                conn.ExecuteQuery();

                string var_Name = conn.GetFieldValue("KET_DESC");

                conn.QueryString = "Select PROG_CODE from APPLICATION where AP_REGNO = '" + regno + "'";
                conn.ExecuteQuery();

                string prog_code = conn.GetFieldValue("PROG_CODE");

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objPaste = null;
                object objCopy = null;
                bool bSukses = true;
                object objValue = null;
                object objType = Type.Missing;
                int iItem = 0;
                int iItemOther = 0;
                int iItemPosition = 0;
                int m_Row = 0;


                System.Data.DataTable dt_field = null;

                conn.QueryString = "Select * from NOTA_ANALISA where NOTA_ID = '" + var_idExport1 + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    string path = conn.GetFieldValue("NOTA_PATH");
                    string file_xls = nota + ".XLSX";
                    string url = conn.GetFieldValue("NOTA_URL");

                    Microsoft.Office.Interop.Excel.Application excelApp = null;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                    Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);

                    try
                    {
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;

                        Process[] newProcess = Process.GetProcessesByName("EXCEL");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        /// Save process into database
                        /// 
                        //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);

                        fileIn = path + file_xls;

                        excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                            false, false, 0, true);

                        excelSheet = excelWorkBook.Worksheets;

                        Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                        fileNm = regno + "-" + nota + "-" + var_user + ".XLSX";
                        fileOut = path + fileNm;
                        //fileResult = url + fileNm;

                        #region Step Fill Multi Ketentuan Kredit
                        // Step 2.1

                        conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_0 where STEP = 1 AND NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD6 '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][3].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][4]);
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][5].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }


                        // sTE 2.2
                        conn.QueryString = "Select NOTA_ID, SEQ, STEP, NOTA_COL, NOTA_ROW, NOTA_FIELD, [DESCRIPTION] from NOTA_ANALISA_DETAIL1_0 where STEP = 0 AND NOTA_ID = '" + nota + "' order by NOTA_ID, NOTA_ROW, SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_NOTA_ANALISA_WORD6 '" + regno + "'";
                        conn.ExecuteQuery();

                        if ((conn.GetRowCount() > 1) && (dt_field.Rows.Count > 1))
                        {
                            iItemPosition = Convert.ToInt32(dt_field.Rows[0][4]);

                            //Prepare Empty Row with all Formats
                            for (int j = 0; j < conn.GetRowCount() - 1; j++)
                            {
                                Microsoft.Office.Interop.Excel.Range excelRange = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "IV" + iItemPosition.ToString().Trim()).EntireRow;
                                excelRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
                                iItemOther = iItemPosition + j;
                            }

                            Microsoft.Office.Interop.Excel.Range exlCopy = excelWorkSheet.get_Range("A" + (iItemOther + 1).ToString().Trim(), "IV" + (iItemOther + 1).ToString().Trim());
                            objCopy = exlCopy.Copy(objType);

                            Microsoft.Office.Interop.Excel.Range exlPaste = excelWorkSheet.get_Range("A" + iItemPosition.ToString().Trim(), "A" + iItemOther.ToString().Trim());
                            objPaste = exlPaste.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        }

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            if (j > 0) m_Row = m_Row + 1;

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                string Col = dt_field.Rows[i][3].ToString().Trim();
                                int Row = Convert.ToInt32(dt_field.Rows[i][4]) + m_Row;
                                string Cell = Col + Row.ToString().Trim();
                                string Field = dt_field.Rows[i][5].ToString();

                                if (Field == "NUM")
                                    objValue = j + 1;
                                else
                                    objValue = conn.GetFieldValue(j, Field);

                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)excelWorkSheet.get_Range(Cell, Cell);
                                excelCell.Value2 = objValue;

                                iItem++;
                            }
                        }
                        #endregion


                        if (iItem > 0)
                        {
                            excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);

                            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                            bSukses = true;
                        }
                        else
                            bSukses = false;

                        if (bSukses)
                        {
                            // Maintenance Table Nota_Export

                            if (var_idExport2 == string.Empty)
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','NOTA','" + fileNm + "','" + userid + "', '1'";
                            else
                                conn.QueryString = "exec CP_NOTA_EXPORT '" + nota + "','" + regno + "','" + var_idExport2 + "','" + fileNm + "', '" + userid + "', '1'";

                            conn.ExecuteQuery();

                            mStatus = "Export Succesfully";

                        }
                        else
                        {
                            mStatus = "No Data to Export";
                        }
                    }
                    catch (Exception e)
                    {
                        //Response.Write("<!-- " + e.Message + "\n" + e.StackTrace + " -->\n");
                        mStatus = e.Message;
                    }
                    finally
                    {
                        if (excelWorkBook != null)
                        {
                            excelWorkBook.Close(true, fileOut, null);
                            excelWorkBook = null;
                        }
                        if (excelApp != null)
                        {
                            excelApp.Workbooks.Close();
                            excelApp.Application.Quit();
                            excelApp = null;
                        }
                    }
                    try
                    {
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }

                    }
                    catch (Exception e)
                    {
                        mStatus = e.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.SPPKExportASPXCreateSPPKWord(string regno, string userid, out Dictionary<string, string> result)
        {
            string mStatus = string.Empty;

            try
            {
                Dictionary<string, string> resultD = new Dictionary<string, string>();

                //string szUser = ddl_manual.SelectedValue;
                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                string mNotaNumber = string.Empty;
                object objType = Type.Missing;
                bool bSukses = true;
                object objValue = null;
                object Bookmark;
                string Field;
                string strObject;
                object objFileIn;
                object objFileOut;
                int bookmark_cnt = 0;
                string vAPP_ROOT = "";

                string LBL_STATUSEXPORT = "LBL_STATUSEXPORT";
                string LBL_STATUS_EXPORT = "LBL_STATUS_EXPORT";

                string LBL_STATUSEXPORT_VAL = "";
                string LBL_STATUS_EXPORT_VAL = "";

                /*System.Web.UI.WebControls.Label LBL_STATUSEXPORT = (System.Web.UI.WebControls.Label)thisPage.FindControl("LBL_STATUSEXPORT");
                System.Web.UI.WebControls.Label LBL_STATUS_EXPORT = (System.Web.UI.WebControls.Label)thisPage.FindControl("LBL_STATUS_EXPORT");
                System.Web.UI.WebControls.Label LBL_STATUS = (System.Web.UI.WebControls.Label)thisPage.FindControl("LBL_STATUS");*/

                conn.QueryString = "select p.businessunit from application a " +
                    " left join rfprogram p on a.prog_code = p.programid " +
                    " where ap_regno = '" + regno + "'";
                conn.ExecuteQuery();

                string business_unit = "";
                try
                {
                    business_unit = conn.GetFieldValue("businessunit");
                }
                catch (Exception e)
                {
                    mStatus = e.Message;
                    result = resultD;
                    return mStatus;
                }


                System.Data.DataTable dt_field = null;
                System.Data.DataTable dt_field1 = null;
                System.Data.DataTable dt_field2 = null;
                System.Data.DataTable dt_field3 = null;


                /// Mengambil application root
                /// 
                conn.QueryString = "select APP_ROOT from APP_PARAMETER";
                conn.ExecuteQuery();
                vAPP_ROOT = conn.GetFieldValue("APP_ROOT");


                conn.QueryString = "select * from rfsppk  " +
                    " where b_unit = '" + business_unit + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() > 0)
                {
                    string nota = conn.GetFieldValue("NOTA_ID");
                    string sheet = conn.GetFieldValue("NOTA_SHEET");
                    //string path = conn.GetFieldValue("NOTA_PATH");
                    string path = vAPP_ROOT + conn.GetFieldValue("NOTA_PATH");
                    string file_doc = nota + ".DOCX";
                    //string url = conn.GetFieldValue("NOTA_URL");
                    string template = conn.GetFieldValue("NOTA_TEMPLATE");
                    string b_unit = conn.GetFieldValue("B_UNIT");
                    string drill = conn.GetFieldValue("DRILL");

                    /// Cek apakah file templatenya (input) ada atau tidak
                    /// 
                    if (!File.Exists(template + file_doc))
                    {
                        //GlobalTools.popMessage(thisPage, "File Template tidak ada!");
                        result = resultD;
                        return "File Template tidak ada!";
                    }

                    /// Cek direktori untuk menyimpan file hasil export (output)
                    /// 
                    if (!Directory.Exists(path))
                    {
                        // create directory if does not exist
                        Directory.CreateDirectory(path);
                    }


                    int iItem = 0;

                    object oMissingObject = System.Reflection.Missing.Value;

                    Microsoft.Office.Interop.Word.Application wordApp = null;
                    Microsoft.Office.Interop.Word.Document wordDoc = null;

                    ArrayList orgId = new ArrayList();
                    ArrayList newId = new ArrayList();

                    //Collecting Existing Winword in Taskbar

                    Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                    foreach (Process thisProcess in oldProcess)
                        orgId.Add(thisProcess);


                    try
                    {
                        wordApp = new Application();
                        wordApp.Visible = false;

                        //Collecting Existing Winword in Taskbar 

                        Process[] newProcess = Process.GetProcessesByName("WINWORD");
                        foreach (Process thisProcess in newProcess)
                            newId.Add(thisProcess);

                        //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);


                        iItem = 0;

                        string var_user = "";
                        fileNm = regno + "-" + nota + "-" + var_user + ".DOCX";

                        //objFileIn = path + file_doc;		
                        //fileResult = url + fileNm;

                        objFileIn = template + file_doc;
                        objFileOut = path + fileNm;

                        //wordDoc = wordApp.Documents.Open(ref objFileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, 
                        //	ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordDoc = wordApp.Documents.Add(ref objFileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                        wordDoc.Activate();

                    }
                    catch
                    {
                        //Response.Write("<!-- " + e1.Message.Replace("'", "") + " -->");
                        /*LBL_STATUS_EXPORT.ForeColor = Color.Red;
                        LBL_STATUSEXPORT.ForeColor = Color.Red;
                        LBL_STATUS_EXPORT.Text = "Fail in creating the Objects";
                        LBL_STATUSEXPORT.Text = "";*/
                        LBL_STATUS_EXPORT_VAL = "Fail in creating the Objects";
                        LBL_STATUSEXPORT_VAL = "";

                        resultD.Add(LBL_STATUS_EXPORT, LBL_STATUS_EXPORT_VAL);
                        resultD.Add(LBL_STATUSEXPORT, LBL_STATUSEXPORT_VAL);

                        try
                        {
                            if (wordDoc != null)
                            {
                                ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                                wordDoc = null;
                            }
                            if (wordApp != null)
                            {
                                ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                                wordApp = null;
                                // Killing Proses after Export
                                for (int x = 0; x < newId.Count; x++)
                                {
                                    Process xnewId = (Process)newId[x];
                                    bool bSameId = false;
                                    for (int z = 0; z < orgId.Count; z++)
                                    {
                                        Process xoldId = (Process)orgId[z];

                                        if (xnewId.Id == xoldId.Id)
                                        {
                                            bSameId = true;
                                            break;
                                        }
                                    }
                                    if (bSameId)
                                    {
                                        try
                                        {
                                            xnewId.Kill();
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                        result = resultD;
                        return "Failed";
                    }

                    Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;
                    Microsoft.Office.Interop.Word.Bookmark oBook;

                    /*  SPPK compose of the followings:
                         *	1.General Info -  ( 1 time ), Category 1
                         *  2.Ketentuan Kredit ( Multiples ... depending .... )
                         *		2.1 Collaterals - depending on Ketentuan
                         *  3.Syarats ...
                         */


                    #region Step Fill General Info
                    // Step 1 - mask out the old codes .... dangerous 
                    //conn.QueryString = "Select * from NOTA_ANALISA_DETAIL where NOTA_ID = '" + nota + "' order by NOTA_ID, SEQ";
                    conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 1  order by SEQ";
                    conn.ExecuteQuery();

                    dt_field = conn.GetDataTable().Copy();

                    conn.QueryString = "exec CP_EXPORT_SPPK_GENERAL '" + regno + "','" + business_unit + "'";
                    conn.ExecuteQuery();

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {
                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {
                            try
                            {
                                Bookmark = dt_field.Rows[i][6];
                                Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(Field);

                                strObject = objValue.ToString();

                                if (wordBookMark.Exists(Bookmark.ToString()))
                                {
                                    //strObject = Convert.ToDateTime(objValue).ToShortDateString();

                                    //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                    oBook = wordBookMark.get_Item(ref Bookmark);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                            catch (Exception e2)
                            {
                                //Response.Write("<!-- " + e2.Message.Replace("'", "") + " -->");
                                mStatus = e2.Message;
                            }
                        }
                    }

                    #endregion

                    #region Step Fill in Ketentuan Kredit

                    conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 2  order by SEQ";
                    conn.ExecuteQuery();

                    dt_field = conn.GetDataTable().Copy();


                    // collaterals list
                    conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 3  order by SEQ";
                    conn.ExecuteQuery();

                    dt_field2 = conn.GetDataTable().Copy();

                    conn.QueryString = "select  distinct kk.ket_code from approval_decision ad " +
                        "left join custproduct cp on ad.productid = cp.productid  and ad.ap_regno = cp.ap_regno and cp.apptype = ad.apptype and cp.prod_seq = ad.prod_seq " +
                        "left join ketentuan_kredit kk on kk.ket_code = cp.ket_code " +
                        "where cp.cp_reject <> 1 AND Ad.ap_regno = '" + regno + "' " +
                        "and ad.ad_seq = (select max(ad_seq) from approval_decision where ap_regno = '" + regno + "')";
                    conn.ExecuteQuery();

                    dt_field1 = conn.GetDataTable().Copy();


                    for (int k = 0; k < dt_field1.Rows.Count; k++)
                    {

                        conn.QueryString = "exec CP_EXPORT_SPPK_KETENTUAN '" + regno + "','" + dt_field1.Rows[k][0] + "'";
                        conn.ExecuteQuery();
                        dt_field3 = conn.GetDataTable().Copy();



                        for (int j = 0; j < dt_field3.Rows.Count; j++)  // could be multiple ketentuan ....
                        {

                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                bookmark_cnt = i;
                                try
                                {
                                    Field = dt_field.Rows[i][4].ToString();

                                    Bookmark = (string)(dt_field.Rows[i][6] + k.ToString());

                                    objValue = dt_field3.Rows[j][Field];

                                    strObject = objValue.ToString();

                                    if (wordBookMark.Exists(Bookmark.ToString()))
                                    {
                                        //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                        oBook = wordBookMark.get_Item(ref Bookmark);
                                        oBook.Select();
                                        oBook.Range.Text = strObject;
                                    }
                                    iItem++;
                                }
                                catch (Exception e3)
                                {
                                    //Response.Write("<!-- " + e3.Message.Replace("'", "") + " -->");
                                    mStatus = e3.Message;
                                }
                            }  // i end
                        } // j end



                        conn.QueryString = "exec CP_EXPORT_SPPK_COLLATERALS '" + regno + "','" + dt_field1.Rows[k][0] + "'";
                        conn.ExecuteQuery();


                        for (int i = 0; i < dt_field2.Rows.Count; i++)
                        {

                            try
                            {
                                //Bookmark = dt_field2.Rows[bookmark_cnt][6];
                                Bookmark = (string)(dt_field2.Rows[i][6] + k.ToString());
                                Field = dt_field2.Rows[i][4].ToString();

                                for (int j = 0; j < conn.GetRowCount(); j++)
                                {
                                    objValue = conn.GetFieldValue(j, Field);
                                    strObject = objValue.ToString() + "\n";

                                    if (wordBookMark.Exists(Bookmark.ToString()))
                                    {
                                        //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                        oBook = wordBookMark.get_Item(ref Bookmark);
                                        oBook.Select();
                                        oBook.Range.Text = strObject;
                                    }
                                    iItem++;
                                } // end of j
                            }
                            catch (Exception e4)
                            {
                                //Response.Write("<!-- " + e4.Message.Replace("'", "") + " -->");
                                mStatus = e4.Message;
                            }
                        } // end of i						
                    }


                    #endregion

                    #region Syarat-syarats
                    // SYARAT TandaTangan
                    conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 4  order by SEQ";
                    conn.ExecuteQuery();

                    dt_field = conn.GetDataTable().Copy();

                    conn.QueryString = "exec CP_EXPORT_SPPK_SYARAT_TANDATANGAN '" + regno + "'";
                    conn.ExecuteQuery();

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {
                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {
                            try
                            {
                                Bookmark = dt_field.Rows[i][6];
                                Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                strObject = objValue.ToString() + "\n";

                                if (wordBookMark.Exists(Bookmark.ToString()))
                                {
                                    //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                    oBook = wordBookMark.get_Item(ref Bookmark);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                            catch (Exception e5)
                            {
                                //Response.Write("<!-- " + e5.Message.Replace("'", "") + " -->");
                                mStatus = e5.Message;
                            }
                        }
                    }

                    // SYARAT PENRARIKAN KREDIT CL
                    conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 5  order by SEQ";
                    conn.ExecuteQuery();

                    dt_field = conn.GetDataTable().Copy();

                    conn.QueryString = "exec CP_EXPORT_SPPK_SYARAT_PENARIKAN_CL '" + regno + "'";
                    conn.ExecuteQuery();

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {
                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {
                            try
                            {
                                Bookmark = dt_field.Rows[i][6];
                                Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                strObject = objValue.ToString() + "\n";

                                if (wordBookMark.Exists(Bookmark.ToString()))
                                {
                                    //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                    oBook = wordBookMark.get_Item(ref Bookmark);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                            catch (Exception e6)
                            {
                                //Response.Write("<!-- " + e6.Message.Replace("'", "") + " -->");
                                mStatus = e6.Message;
                            }
                        }
                    }

                    if (business_unit != "SM100")
                    {
                        // SYARAT PENRARIKAN KREDIT NCL
                        conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 6  order by SEQ";
                        conn.ExecuteQuery();

                        dt_field = conn.GetDataTable().Copy();

                        conn.QueryString = "exec CP_EXPORT_SPPK_SYARAT_PENARIKAN_NCL '" + regno + "'";
                        conn.ExecuteQuery();

                        for (int j = 0; j < conn.GetRowCount(); j++)
                        {
                            for (int i = 0; i < dt_field.Rows.Count; i++)
                            {
                                try
                                {
                                    Bookmark = dt_field.Rows[i][6];
                                    Field = dt_field.Rows[i][4].ToString();

                                    objValue = conn.GetFieldValue(j, Field);

                                    strObject = objValue.ToString() + "\n";

                                    if (wordBookMark.Exists(Bookmark.ToString()))
                                    {
                                        //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                        oBook = wordBookMark.get_Item(ref Bookmark);
                                        oBook.Select();
                                        oBook.Range.Text = strObject;
                                    }
                                    iItem++;
                                }
                                catch (Exception e7)
                                {
                                    //Response.Write("<!-- " + e7.Message.Replace("'", "") + " -->");
                                    mStatus = e7.Message;
                                }
                            }
                        }

                    }

                    // SYARAT syarat lain
                    conn.QueryString = "Select NOTA_ID,SEQ,NOTA_COL,NOTA_ROW,NOTA_FIELD,[DESCRIPTION],BOOKMARK from rfsppkdetail where NOTA_ID = '" + nota + "' and category = 7  order by SEQ";
                    conn.ExecuteQuery();

                    dt_field = conn.GetDataTable().Copy();

                    conn.QueryString = "exec CP_EXPORT_SPPK_SYARAT_LAIN2 '" + regno + "'";
                    conn.ExecuteQuery();

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {
                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {
                            try
                            {
                                Bookmark = dt_field.Rows[i][6];
                                Field = dt_field.Rows[i][4].ToString();

                                objValue = conn.GetFieldValue(j, Field);

                                strObject = objValue.ToString() + "\n";

                                if (wordBookMark.Exists(Bookmark.ToString()))
                                {
                                    //Word.Bookmark oBook = wordBookMark.Item(ref Cell);
                                    oBook = wordBookMark.get_Item(ref Bookmark);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                                iItem++;
                            }
                            catch (Exception e8)
                            {
                                //Response.Write("<!-- " + e8.Message.Replace("'", "") + " -->");
                                mStatus = e8.Message;
                            }
                        }
                    }
                    #endregion


                    ///* start -- simpen hasil export					

                    if (iItem > 0)
                    {
                        wordDoc.SaveAs(ref objFileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                            ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));
                        bSukses = true;
                    }
                    else
                        bSukses = false;

                    if (bSukses)
                    {
                        // Maintenance Table SPPK_Export

                        conn.QueryString = "exec CP_SPPK_EXPORT '" + nota + "','" + regno + "','" + fileNm + "','','" + userid + "', '1'";
                        //conn.QueryString = "exec CP_SPPK_EXPORT '" + nota +"','" + Request.QueryString["regno"] + "','" + var_idExport2 + "','" + fileNm + "','', '" + Session["UserID"] + "', '1'";

                        conn.ExecuteQuery();
                        mStatus = "Export Succesfully";

                    }
                    else
                    {
                        mStatus = "No Data to Export";
                    }

                    if (wordDoc != null)
                    {
                        ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordDoc = null;
                    }
                    if (wordApp != null)
                    {
                        ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordApp = null;

                        // Killing Proses after Export
                        for (int x = 0; x < newId.Count; x++)
                        {
                            Process xnewId = (Process)newId[x];

                            bool bSameId = false;
                            for (int z = 0; z < orgId.Count; z++)
                            {
                                Process xoldId = (Process)orgId[z];

                                if (xnewId.Id == xoldId.Id)
                                {
                                    bSameId = true;
                                    break;
                                }
                            }
                            if (bSameId)
                            {
                                try
                                {
                                    xnewId.Kill();
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }

                    }



                    //ViewFileExport();
                    /*
                    catch (Exception ex) {

                        LBL_STATUS_EXPORT.ForeColor = Color.Red;
                        LBL_STATUSEXPORT.ForeColor = Color.Red;

                        LBL_STATUS_EXPORT.Text = "Error Exporting File!";
                        LBL_STATUSEXPORT.Text = ex.ToString();
                    }
                    */

                }
                result = resultD;
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CustomerInfoExportASPXExport_Excel(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                System.Data.DataTable dt_field = null;
                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objType = Type.Missing;
                
                string var_idExport = DDL_FORMATFILESelectedValue;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                conn.QueryString = "select APP_ROOT from APP_PARAMETER";
                conn.ExecuteQuery();
                string vAPP_ROOT = conn.GetFieldValue("APP_ROOT");


                /// Mengambil nilai parameter
                /// 
                conn.QueryString = " select * from RFCUSTEXPORT where EXPORT_ID = '" + var_idExport + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() == 0)
                {
                    //GlobalTools.popMessage(thisPage, "Data Referensi RFCUSTEXPORT kosong!");
                    return "Data Referensi RFCUSTEXPORT kosong!";
                }

                string nota = var_idExport;										// nama file hasil export
                string sheet = conn.GetFieldValue("EXPORT_SHEET");			// sheet di excel
                string path = vAPP_ROOT + conn.GetFieldValue("EXPORT_PATH");	// directory excel hasil export			
                string file_xls = nota + ".XLSX";							// nama file excel template
                string template = conn.GetFieldValue("EXPORT_TEMPLATE");		// directory excel template
                string url = conn.GetFieldValue("EXPORT_URL");				// url (link) untuk download			
                //string procedure_name = conn.GetFieldValue ("STOREPROCEDURE");



                /// Men-construct nama file
                /// 
                fileIn = template + file_xls;	// file template
                fileNm = curef + "-" + nota + "-" + userid + ".XLSX";	// file hasil export
                fileOut = path + fileNm;


                /// Cek apakah file templatenya (input) ada atau tidak
                /// 
                if (!File.Exists(template + file_xls))
                {
                    //GlobalTools.popMessage(thisPage, "File Template tidak ada!");
                    return "File Template tidak ada!";
                }

                /// Cek direktori untuk menyimpan file hasil export (output)
                /// 
                if (!Directory.Exists(path))
                {
                    // create directory if does not exist
                    Directory.CreateDirectory(path);
                }


                /// dapatkan semua fields to populate
                /// 

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);


                // Always already when using Export Excel file format					
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");


                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                Process[] newProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in newProcess)
                    newId.Add(thisProcess);

                /// Save process into database
                /// 
                //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);


                excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                    false, false, 0, true);

                excelSheet = excelWorkBook.Worksheets;

                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                #region " Fill Customer Information "
                try
                {
                    conn.QueryString = "Select SEQ, EXPORT_COL, EXPORT_ROW, EXPORT_FIELD, [DESCRIPTION] from RFCUSTEXPORTDETAIL" +
                        " where EXPORT_ID = '" + nota + "' order by SEQ";
                    conn.ExecuteQuery();
                    dt_field = conn.GetDataTable().Copy();


                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }
                #endregion


                try
                {
                    /// Save file fisik hasil export
                    /// 
                    //excelWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));


                    /// Save data file hasil export ke database
                    /// 
                    conn.QueryString = "exec IN_CUST_EXPORT '" +
                        var_idExport + "','" + curef + "', '" +
                        fileNm + "', '" +
                        userid + " ', '1'";
                    conn.ExecuteNonQuery();
                    mStatus = "Export Succesfully";
                }
                catch { }
                //}



                /// Kill Process
                /// 
                try
                {
                    // close the excel objects
                    if (excelWorkBook != null)
                    {
                        excelWorkBook.Close(true, fileOut, null);
                        excelWorkBook = null;
                    }

                    if (excelApp != null)
                    {
                        excelApp.Workbooks.Close();
                        excelApp.Application.Quit();
                        excelApp = null;
                    }
                }
                catch { }

                try
                {
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }

                }
                catch { }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }
            return mStatus;

        }

        string IWord.CustomerInfoExportASPXExport_Word(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                System.Data.DataTable dt_field = null;
                System.Data.DataTable dt_proc = null;

                string var_idExport = DDL_FORMATFILESelectedValue;

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objValue = null;
                object objType = Type.Missing;
            
                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                /// Mengambil application root
                /// 
                conn.QueryString = "select APP_ROOT from APP_PARAMETER";
                conn.ExecuteQuery();
                string vAPP_ROOT = conn.GetFieldValue("APP_ROOT");

                /// Mengambil nilai parameter
                /// 
                conn.QueryString = " select * from RFCUSTEXPORT where EXPORT_ID = '" + var_idExport + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() == 0)
                {
                    //GlobalTools.popMessage(thisPage, "Data Referensi RFCUSTEXPORT kosong!");
                    return "Data Referensi RFCUSTEXPORT kosong!";
                }


                string nota = var_idExport;											// nama file hasil export
                string path = vAPP_ROOT + conn.GetFieldValue("EXPORT_PATH");		// path untuk upload
                string file_xls = nota + ".DOCX";									// nama file word template
                string template = conn.GetFieldValue("EXPORT_ID");					// nama word template
                string template_path = conn.GetFieldValue("EXPORT_TEMPLATE");		// directory word template
                string url = conn.GetFieldValue("EXPORT_URL");						// url (link) untuk download


                fileNm = curef + "-" + nota + "-" + userid + ".DOCX";

                object objFileIn = template_path + file_xls;
                object objFileOut = path + fileNm;

                ////////////////////////////////////////////////////////////////////////////
                /// Cek apakah file templatenya (input) ada atau tidak
                /// 
                if (!File.Exists(template_path + file_xls))
                {
                    //GlobalTools.popMessage(thisPage, "File Template tidak ada!");
                    return "File Template tidak ada!";
                }


                /////////////////////////////////////////////////////////////////////////////
                /// Cek direktori untuk menyimpan file hasil export (output)
                /// 
                if (!Directory.Exists(path))
                {
                    // create directory if does not exist
                    Directory.CreateDirectory(path);
                }


                object oMissingObject = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document wordDoc = null;

                Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);


                /// 
                /// Always already when using Export Excel file format					
                /// 
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");


                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = false;

                /// Collecting Existing Winword in Taskbar 
                /// 
                Process[] newProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in newProcess)
                    newId.Add(thisProcess);

                /// 
                /// Save word process into database
                /// 					
                try
                {
                    //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);
                }
                catch (Exception ex)
                {
                    mStatus = ex.Message;
                }


                wordDoc = wordApp.Documents.Open(ref objFileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                    ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);
                wordDoc.Activate();
                Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;


                object oCell;
                string tempField;
                object sField;
                string strObject;

                conn.QueryString = "select * from RFCUSTEXPORTPROC where EXPORT_ID = '" + var_idExport + "'";
                conn.ExecuteQuery();
                dt_proc = conn.GetDataTable().Copy();


                //////////////////////////////////////////////////////////////////////////
                /// Populate data to word file using different stored procedure
                /// 
                for (int p = 0; p < dt_proc.Rows.Count; p++)
                {
                    string storedproc = dt_proc.Rows[p]["STOREDPROCEDURE"].ToString();

                    /// if no stored procedure defined, continue to the next one
                    /// 
                    if (storedproc.Length == 0) continue;

                    /// Get fields from db to map to bookmark in word file
                    /// 
                    conn.QueryString = "select d.export_id, d.seq, d.export_col, d.export_row, d.export_field, " +
                                        " d.[description], d.[group], d.category, p.storedprocedure " +
                                        " from rfcustexportdetail d " +
                                        " left join rfcustexportproc p on d.export_id = p.export_id and d.category = p.seq " +
                                        " where d.export_id = '" + var_idExport + "' and p.storedprocedure = '" + storedproc + "' " +
                                        " order by d.seq";
                    conn.ExecuteQuery();
                    dt_field = conn.GetDataTable().Copy();


                    /// Execute each stored procedure
                    /// 
                    try
                    {
                        conn.QueryString = "exec " + storedproc + " '" + curef + "', '" + regno + "'";
                        conn.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        /// There's no such stored procedure in db ??
                        /// 
                        String a = ex.Message;
                        //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                        continue;
                    }

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {
                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {
                            try
                            {
                                oCell = dt_field.Rows[i]["export_col"];
                                tempField = dt_field.Rows[i]["export_field"].ToString();
                                sField = dt_field.Rows[i]["export_field"].ToString();

                                objValue = conn.GetFieldValue(j, tempField);

                                if (wordBookMark.Exists(sField.ToString()))
                                {

                                    if (dt_field.Rows[i]["Group"].ToString() != "0") strObject = objValue.ToString();
                                    else strObject = objValue.ToString() + "\n";

                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref sField);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                            }
                            catch { }

                        } //endloop var i
                    } //endloop var j
                }


                ////////////////////////////////////////////////////////////////////////////
                /// Save Word File
                try
                {
                    wordDoc.SaveAs(ref objFileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                        ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));

                    /// 
                    /// Menyimpan data hasil export ke database
                    /// 
                    conn.QueryString = "exec IN_CUST_EXPORT '" + nota + "','" + curef + "','" + fileNm + "', '" + userid + "', '1'";

                    conn.ExecuteQuery();
                    mStatus = "Export Succesfully";
                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }


                // try to close word dulu ...
                try
                {
                    if (wordDoc != null)
                    {
                        ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordDoc = null;
                    }
                    if (wordApp != null)
                    {
                        //wordApp.Application.Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordApp = null;
                    }
                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }

                /// Kill process
                /// 
                try
                {

                    // Killing Proses after Export
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }

                    } // end x		
                }
                catch (Exception ex)
                {
                    mStatus = ex.Message;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CBICustomerInfoExportASPXExport_Excel(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                System.Data.DataTable dt_field = null;
                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objType = Type.Missing;
            
                string var_idExport = DDL_FORMATFILESelectedValue;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                conn.QueryString = "select APP_ROOT from APP_PARAMETER";
                conn.ExecuteQuery();
                string vAPP_ROOT = conn.GetFieldValue("APP_ROOT");


                /// Mengambil nilai parameter
                /// 
                conn.QueryString = " select * from RFCUSTEXPORT where EXPORT_ID = '" + var_idExport + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() == 0)
                {
                    //GlobalTools.popMessage(thisPage, "Data Referensi RFCUSTEXPORT kosong!");
                    return "Data Referensi RFCUSTEXPORT kosong!";
                }

                string nota = var_idExport;										// nama file hasil export
                string sheet = conn.GetFieldValue("EXPORT_SHEET");			// sheet di excel
                string path = vAPP_ROOT + conn.GetFieldValue("EXPORT_PATH");	// directory excel hasil export			
                string file_xls = nota + ".XLSX";							// nama file excel template
                string template = conn.GetFieldValue("EXPORT_TEMPLATE");		// directory excel template
                string url = conn.GetFieldValue("EXPORT_URL");				// url (link) untuk download			
                //string procedure_name = conn.GetFieldValue ("STOREPROCEDURE");



                /// Men-construct nama file
                /// 
                fileIn = template + file_xls;	// file template
                fileNm = curef + "-" + nota + "-" + userid + ".XLSX";	// file hasil export
                fileOut = path + fileNm;


                /// Cek apakah file templatenya (input) ada atau tidak
                /// 
                if (!File.Exists(template + file_xls))
                {
                    //GlobalTools.popMessage(thisPage, "File Template tidak ada!");
                    return "File Template tidak ada!";
                }

                /// Cek direktori untuk menyimpan file hasil export (output)
                /// 
                if (!Directory.Exists(path))
                {
                    // create directory if does not exist
                    Directory.CreateDirectory(path);
                }


                /// dapatkan semua fields to populate
                /// 

                Microsoft.Office.Interop.Excel.Application excelApp = null;
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
                Microsoft.Office.Interop.Excel.Sheets excelSheet = null;

                Process[] oldProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);


                // Always already when using Export Excel file format					
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");


                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                Process[] newProcess = Process.GetProcessesByName("EXCEL");
                foreach (Process thisProcess in newProcess)
                    newId.Add(thisProcess);

                /// Save process into database
                /// 
                //SupportTools.saveProcessExcel(excelApp, newId, orgId, conn);


                excelWorkBook = excelApp.Workbooks.Open(fileIn, 0, false, 5, string.Empty, string.Empty, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t|",
                    false, false, 0, true);

                excelSheet = excelWorkBook.Worksheets;

                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelSheet.get_Item(sheet);

                #region " Fill Customer Information "
                try
                {
                    conn.QueryString = "Select SEQ, EXPORT_COL, EXPORT_ROW, EXPORT_FIELD, [DESCRIPTION] from RFCUSTEXPORTDETAIL" +
                        " where EXPORT_ID = '" + nota + "' order by SEQ";
                    conn.ExecuteQuery();
                    dt_field = conn.GetDataTable().Copy();


                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }
                #endregion


                try
                {
                    /// Save file fisik hasil export
                    /// 
                    //excelWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    excelWorkBook.SaveAs(fileOut, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, null, null,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, true);
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));


                    /// Save data file hasil export ke database
                    /// 
                    conn.QueryString = "exec IN_CUST_EXPORT '" +
                        var_idExport + "','" + curef + "', '" +
                        fileNm + "', '" +
                        userid + " ', '1'";
                    conn.ExecuteNonQuery();
                    mStatus = "Export Succesfully";
                }
                catch { }
                //}



                /// Kill Process
                /// 
                try
                {
                    // close the excel objects
                    if (excelWorkBook != null)
                    {
                        excelWorkBook.Close(true, fileOut, null);
                        excelWorkBook = null;
                    }

                    if (excelApp != null)
                    {
                        excelApp.Workbooks.Close();
                        excelApp.Application.Quit();
                        excelApp = null;
                    }
                }
                catch { }

                try
                {
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }

                }
                catch(Exception e) 
                {
                    mStatus = e.Message;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;
        }

        string IWord.CBICustomerInfoExportASPXExport_Word(string regno, string userid, string curef, string DDL_FORMATFILESelectedValue)
        {
            string mStatus = string.Empty;

            try
            {
                System.Data.DataTable dt_field = null;
                System.Data.DataTable dt_proc = null;

                string fileNm = string.Empty;
                string fileIn = string.Empty;
                string fileOut = string.Empty;
                object objValue = null;
                object objType = Type.Missing;

                string var_idExport = DDL_FORMATFILESelectedValue;

                ArrayList orgId = new ArrayList();
                ArrayList newId = new ArrayList();

                /// Mengambil application root
                /// 
                conn.QueryString = "select APP_ROOT from APP_PARAMETER";
                conn.ExecuteQuery();
                string vAPP_ROOT = conn.GetFieldValue("APP_ROOT");

                /// Mengambil nilai parameter
                /// 
                conn.QueryString = " select * from RFCUSTEXPORT where EXPORT_ID = '" + var_idExport + "'";
                conn.ExecuteQuery();

                if (conn.GetRowCount() == 0)
                {
                    //GlobalTools.popMessage(thisPage, "Data Referensi RFCUSTEXPORT kosong!");
                    return "Data Referensi RFCUSTEXPORT kosong!";
                }


                string nota = var_idExport;											// nama file hasil export
                string path = vAPP_ROOT + conn.GetFieldValue("EXPORT_PATH");		// path untuk upload
                string file_xls = nota + ".DOCX";									// nama file word template
                string template = conn.GetFieldValue("EXPORT_ID");					// nama word template
                string template_path = conn.GetFieldValue("EXPORT_TEMPLATE");		// directory word template
                string url = conn.GetFieldValue("EXPORT_URL");						// url (link) untuk download


                fileNm = curef + "-" + nota + "-" + userid + ".DOCX";

                object objFileIn = template_path + file_xls;
                object objFileOut = path + fileNm;

                ////////////////////////////////////////////////////////////////////////////
                /// Cek apakah file templatenya (input) ada atau tidak
                /// 
                if (!File.Exists(template_path + file_xls))
                {
                    //GlobalTools.popMessage(thisPage, "File Template tidak ada!");
                   return "File Template tidak ada!";
                }


                /////////////////////////////////////////////////////////////////////////////
                /// Cek direktori untuk menyimpan file hasil export (output)
                /// 
                if (!Directory.Exists(path))
                {
                    // create directory if does not exist
                    Directory.CreateDirectory(path);
                }


                object oMissingObject = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document wordDoc = null;

                Process[] oldProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in oldProcess)
                    orgId.Add(thisProcess);


                /// 
                /// Always already when using Export Excel file format					
                /// 
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");


                wordApp = new Application();
                wordApp.Visible = false;

                /// Collecting Existing Winword in Taskbar 
                /// 
                Process[] newProcess = Process.GetProcessesByName("WINWORD");
                foreach (Process thisProcess in newProcess)
                    newId.Add(thisProcess);

                /// 
                /// Save word process into database
                /// 					
                //SupportTools.saveProcessWord(wordApp, newId, orgId, conn);	


                wordDoc = wordApp.Documents.Open(ref objFileIn, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                    ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);
                wordDoc.Activate();
                Microsoft.Office.Interop.Word.Bookmarks wordBookMark = (Microsoft.Office.Interop.Word.Bookmarks)wordDoc.Bookmarks;


                object oCell;
                string tempField;
                object sField;
                string strObject;

                conn.QueryString = "select * from RFCUSTEXPORTPROC where EXPORT_ID = '" + var_idExport + "'";
                conn.ExecuteQuery();
                dt_proc = conn.GetDataTable().Copy();


                //////////////////////////////////////////////////////////////////////////
                /// Populate data to word file using different stored procedure
                /// 
                for (int p = 0; p < dt_proc.Rows.Count; p++)
                {
                    string storedproc = dt_proc.Rows[p]["STOREDPROCEDURE"].ToString();

                    /// if no stored procedure defined, continue to the next one
                    /// 
                    if (storedproc.Length == 0) continue;

                    /// Get fields from db to map to bookmark in word file
                    /// 
                    conn.QueryString = "select d.export_id, d.seq, d.export_col, d.export_row, d.export_field, " +
                                        " d.[description], d.[group], d.category, p.storedprocedure " +
                                        " from rfcustexportdetail d " +
                                        " left join rfcustexportproc p on d.export_id = p.export_id and d.category = p.seq " +
                                        " where d.export_id = '" + var_idExport + "' and p.storedprocedure = '" + storedproc + "' " +
                                        " order by d.seq";
                    conn.ExecuteQuery();
                    dt_field = conn.GetDataTable().Copy();


                    /// Execute each stored procedure
                    /// 
                    try
                    {
                        conn.QueryString = "exec " + storedproc + " '" + curef + "', '" + regno + "'";
                        conn.ExecuteQuery();
                    }
                    catch
                    {
                        /// There's no such stored procedure in db ??
                        /// 
                        //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                        continue;
                    }

                    for (int j = 0; j < conn.GetRowCount(); j++)
                    {
                        for (int i = 0; i < dt_field.Rows.Count; i++)
                        {
                            try
                            {
                                oCell = dt_field.Rows[i]["export_col"];
                                tempField = dt_field.Rows[i]["export_field"].ToString();
                                sField = dt_field.Rows[i]["export_field"].ToString();

                                objValue = conn.GetFieldValue(j, tempField);

                                if (wordBookMark.Exists(sField.ToString()))
                                {

                                    if (dt_field.Rows[i]["Group"].ToString() != "0") strObject = objValue.ToString();
                                    else strObject = objValue.ToString() + "\n";

                                    Microsoft.Office.Interop.Word.Bookmark oBook = wordBookMark.get_Item(ref sField);
                                    oBook.Select();
                                    oBook.Range.Text = strObject;
                                }
                            }
                            catch { }

                        } //endloop var i
                    } //endloop var j
                }


                ////////////////////////////////////////////////////////////////////////////
                /// Save Word File
                try
                {
                    wordDoc.SaveAs(ref objFileOut, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject,
                        ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject, ref oMissingObject);

                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(3));

                    /// 
                    /// Menyimpan data hasil export ke database
                    /// 
                    conn.QueryString = "exec IN_CUST_EXPORT '" + nota + "','" + curef + "','" + fileNm + "', '" + userid + "', '1'";

                    conn.ExecuteQuery();
                    mStatus = "Export Succesfully";
                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }


                // try to close word dulu ...
                try
                {
                    if (wordDoc != null)
                    {
                        ((Microsoft.Office.Interop.Word._Document)wordDoc).Close(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordDoc = null;
                    }
                    if (wordApp != null)
                    {
                        ((Microsoft.Office.Interop.Word._Application)wordApp).Quit(ref oMissingObject, ref oMissingObject, ref oMissingObject);
                        wordApp = null;
                    }
                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }

                /// Kill process
                /// 
                try
                {

                    // Killing Proses after Export
                    for (int x = 0; x < newId.Count; x++)
                    {
                        Process xnewId = (Process)newId[x];

                        bool bSameId = false;
                        for (int z = 0; z < orgId.Count; z++)
                        {
                            Process xoldId = (Process)orgId[z];

                            if (xnewId.Id == xoldId.Id)
                            {
                                bSameId = true;
                                break;
                            }
                        }
                        if (bSameId)
                        {
                            try
                            {
                                xnewId.Kill();
                            }
                            catch
                            {
                                continue;
                            }
                        }

                    } // end x		
                }
                catch (Exception ex)
                {
                    //ExceptionHandling.Handler.saveExceptionIntoWindowsLog(ex, Request.Path, "CU_REF: " + LBL_CUREF.Text);
                    mStatus = ex.Message;
                }
            }
            catch (Exception ex)
            {
                ServiceData myServiceData = new ServiceData();
                myServiceData.Result = false;
                myServiceData.ErrorMessage = "Unforeseen error occured. Please try later.";
                myServiceData.ErrorDetails = ex.ToString();
                throw new FaultException<ServiceData>(myServiceData, ex.ToString());
            }

            return mStatus;

        }
    }
}
