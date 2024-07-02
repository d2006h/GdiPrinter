using System;
using System.Linq;
using System.Xml.Linq;
using System.Printing;
using System.Threading;
using System.Timers;
using System.Reflection;
using System.Globalization;
using GemBox.Spreadsheet;
using System.IO;

namespace AbtgGdiPrinter
{
    public class ExcelClass
    {
        //----------------------------LocalVar---------------------//
        string m_formFileAddress;
        ExcelFile m_objWorkbook;
        ExcelWorksheet m_objWorkSheet;
        private string m_strPrinterName;
        System.Timers.Timer m_timerPrint = new System.Timers.Timer();
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger("GDPrinterLog");
        //-------------------------------------------------------//

        public bool OpenExcelFile()
        {
            //----------------------------LocalVar--------------------//
            bool bRet = false;
            //-------------------------------------------------------//
            try
            {
                using (FileStream stream = File.OpenRead(m_formFileAddress))
                {
                    m_objWorkbook = ExcelFile.Load(m_formFileAddress);
                }
                if (m_objWorkbook != null)
                {
                    m_objWorkSheet = m_objWorkbook.Worksheets[0];
                }
                else
                {
                    bRet = false;
                }

            }
            catch (Exception ex)
            {
                throw;
            }
            
            return bRet;

        }
        // set excell trial licence and save template form file address
        public ExcelClass(string formFileAddress)
        {
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                //SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");//GEM 
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.PARAM, "<->{0}:{1} input Params is :\r\nFileNameAndPath:{2}", GetType().Name, MethodBase.GetCurrentMethod().Name, strFileNameAndPath);
                m_formFileAddress = formFileAddress;
            }
            catch (Exception ex)
            {
                throw;
            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);

        }



        //read xml config file and fill the general properties
        public void SetPrinterConfig(string receiptFormFileAddress)
        {
            //----------------------------LocalVar---------------------//
            XDocument xdoc = null;
            //-------------------------------------------------------//try
            try{
                xdoc = XDocument.Load(PrinterClass.m_strXMLConfigPath);
                if (xdoc != null)
                {}
                else
                {}
                var format = new NumberFormatInfo();
                format.NegativeSign = "-";
                format.NumberDecimalSeparator = ".";
                foreach (XElement element in xdoc.Descendants("Printer"))
                {
                    m_strPrinterName = element.Attribute(DefinitionsClass.STR_PRINTER_NAME).Value;                    
                }
                //change and load margines in form xml file
                xdoc = XDocument.Load(receiptFormFileAddress);
                foreach (var item in xdoc.Descendants("Margines"))
                {
                    DefinitionsClass.LeftMargin = Double.Parse(item.Attribute("LeftMargin").Value, format);
                    DefinitionsClass.RightMargin = Double.Parse(item.Attribute("RightMargin").Value, format);
                    DefinitionsClass.TopMargin = Double.Parse(item.Attribute("TopMargin").Value, format);
                    DefinitionsClass.BottomMargin = Double.Parse(item.Attribute("BottomMargin").Value, format);
                }
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "-->{0}:{1} Printer name is -->  {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, m_strPrinterName);

            }
            catch (Exception ex)
            {
                throw;
            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);


        }
        public void PrintMyExcelFile(string strFileName)
        {
            //----------------------------LocalVar---------------------//

            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.PARAM, "<->{0}:{1} input Params is :.....{2}", GetType().Name, MethodBase.GetCurrentMethod().Name, strFileName);

                logger.Debug("set print configurations ");
                PrintOptions printOptions = new PrintOptions();
                printOptions.FromPage = 1;
                printOptions.ToPage = 1;
                printOptions.CopyCount = 1;
                printOptions.DocumentName = m_strPrinterName;
                printOptions.MetafileScaleFactor = 1;
                printOptions.PagesPerSheet = 1;
                printOptions.SelectionType = SelectionType.ActiveSheet;

                logger.Debug("Set page configurations");
                m_objWorkSheet.PrintOptions.HorizontalCentered = true;
                m_objWorkSheet.PrintOptions.TopMargin = DefinitionsClass.TopMargin;
                m_objWorkSheet.PrintOptions.RightMargin = DefinitionsClass.RightMargin;
                m_objWorkSheet.PrintOptions.LeftMargin = DefinitionsClass.LeftMargin;
                m_objWorkSheet.PrintOptions.BottomMargin = DefinitionsClass.BottomMargin;
                m_objWorkSheet.PrintOptions.PrintBlackWhite = true;
                m_objWorkSheet.PrintOptions.FitWorksheetWidthToPages = 1;
                m_objWorkSheet.PrintOptions.FitToPage = true;

                //m_objWorkSheet.PrintOptions.PaperType = PaperType.Custom;
                //Save latest Receipt data in disk
                logger.Debug("Print PDF");
                m_objWorkbook.Save("lastestReceipt1.PDF");
                logger.Debug("Print jpg");
                m_objWorkbook.Save("lastestReceipt2.jpeg");
                try
                {
                    //full
                    /*logger.Debug($"Print the page to printer: {m_strPrinterName} with the print options");
                    m_objWorkbook.Print(m_strPrinterName, printOptions);*/
                    // no options
                    /*logger.Debug($"Print the page to printer: {m_strPrinterName} with no settings");
                    m_objWorkbook.Print(m_strPrinterName);*/
                    // no name 
                    logger.Debug($"Print the page to printer: {m_strPrinterName} with no settings");
                    m_objWorkbook.Print();
                    // wait for job finish
                    CheckIsPrinting(m_strPrinterName);
                    // Cleanup:
                    logger.Debug($"clean UP");
                    GC.Collect();
                    logger.Debug($"Wait for pending job to finish");
                    GC.WaitForPendingFinalizers();
                    logger.Debug("Job Done");
                }
                catch (Exception ex)
                {
                    logger.Fatal(ex.Message);
                    logger.Fatal(ex.Data);
                    logger.Fatal(ex);
                }
                logger.Debug($"check the printer is on job {m_strPrinterName}");
                /*CheckIsPrinting(m_strPrinterName);
                // Cleanup:
                logger.Debug($"clean UP");
                GC.Collect();
                logger.Debug($"Wait for pending job to finish");
                GC.WaitForPendingFinalizers();
                logger.Debug("Job Done");*/

            }
            catch (Exception ex)
            {
                logger.Fatal(ex.Message);
                logger.Fatal(ex);
            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);
        }
        private void CheckIsPrinting(string m_strPrinterName)
        {
            if (PrinterClass.g_CheckIsprintingStatus)
            {
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "<->{0}:{1} m_boolCheckIsprintingStatus is True ", GetType().Name, MethodBase.GetCurrentMethod().Name);

                //----------------------------LocalVar---------------------//
                PrintQueueCollection printQueues = null;
                var server = new LocalPrintServer();
                try
                {
                    //m_objTrace.Trace(TraceLog.TRC_LEVEL.PARAM, "<->{0}:{1} input Params is :.....{2}", GetType().Name, MethodBase.GetCurrentMethod().Name, m_strPrinterName);

                    m_timerPrint.Elapsed += new ElapsedEventHandler(PrinterFailedToPrint);

                    m_timerPrint.Interval = PrinterClass.g_waitingUntillPriinterStartJob;
                    m_timerPrint.Enabled = true;
                    m_timerPrint.Start();

                    PrintQueue queue = server.DefaultPrintQueue;
                    printQueues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });

                    foreach (PrintQueue printQueue in printQueues)
                    {
                        if (printQueue.Name == m_strPrinterName)
                        {
                            while (m_timerPrint.Enabled)
                            {
                                if (printQueue.IsPrinting)
                                {
                                    PrinterClass.g_SuccessfulPrint = true;
                                    //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "<->{0}:{1} Print Has Sent to Printer Successfuly", GetType().Name, MethodBase.GetCurrentMethod().Name);
                                    break;
                                }
                                Thread.Sleep(100);
                            }
                            if (PrinterClass.g_SuccessfulPrint == false)
                            {
                                printQueue.Purge();
                                //     using (PrintServer ps = new PrintServer())
                                //  {
                                //      using (PrintQueue pq = new PrintQueue(ps, printQueue.Name,
                                //            PrintSystemDesiredAccess.AdministratePrinter))
                                //      {
                                //          pq.Purge();
                                //         //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "<->{0}:{1} Print JOBS Successfuly Deleted", GetType().Name, MethodBase.GetCurrentMethod().Name);

                                //     }
                                // }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);
            }
            else
            {
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "<->{0}:{1} m_boolCheckIsprintingStatus is False ", GetType().Name, MethodBase.GetCurrentMethod().Name);

                PrinterClass.g_SuccessfulPrint = true;
            }


        }
        private void PrinterFailedToPrint(object sender, ElapsedEventArgs e)
        {
            //----------------------------LocalVar---------------------//
            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                m_timerPrint.Stop();
                m_timerPrint.Enabled = false;
                PrinterClass.g_SuccessfulPrint = false;
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "<->{0}:{1} Printer Has Some Error", GetType().Name, MethodBase.GetCurrentMethod().Name);
            }

            catch (Exception ex)
            {
                throw;
            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);



        }
        internal void WriteToCell(string[,] CellItems)   // GEM
        {
            //----------------------------LocalVar---------------------//
            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.PARAM, "<->{0}:{1} Input Params: ...{2} ", GetType().Name, MethodBase.GetCurrentMethod().Name, CellItems);
                var style = new CellStyle();

                for (int i = 0; i < CellItems.GetLength(0); i++)
                {
                    m_objWorkSheet.Cells[CellItems[i, 1]].Value = CellItems[i, 4];
                    style.Font.Name = CellItems[i, 0];
                    style.Font.Size = Convert.ToInt32(CellItems[i, 3]);
                    style.Font.Weight = ExcelFont.BoldWeight;
                }
                m_objWorkbook.Save(m_formFileAddress);
            }
            catch (Exception ex)
            {
                throw;
            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);

        }
    }
}

