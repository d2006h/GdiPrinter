using System;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace AbtgGdiPrinter
{
    public class PrinterClass
    {
        private static readonly log4net.ILog logger = log4net.LogManager.GetLogger("GDPrinterLog");

        ExcelClass m_objExcelClass;
        
        public static bool g_SuccessfulPrint = false;
        public static string m_strXMLConfigPath = string.Empty;
        public static bool g_CheckIsprintingStatus = true;
        public static int g_waitingUntillPriinterStartJob = 15000;

        string[,] CellItems;
        int m_intLastRow = 0;
        // read configuration parameters and save the file address
        public PrinterClass(string strXMLConfigPath)
        {
            try
            {
                logger.InfoFormat($"Printer class start {strXMLConfigPath}");
                CheckIsPrintingFlag();
                m_strXMLConfigPath = strXMLConfigPath;

            }
            catch (Exception e)
            {
                logger.Fatal(e.Message);
            }
        }
        // make xdoc and find the xml configuration
        // then it will be check if the print succesfully moved to Printer spool
        private void CheckIsPrintingFlag()
        {
            //----------------------------LocalVar---------------------//
            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                //----------------------------LocalVar---------------------//
                XDocument xdoc = null;
                //-------------------------------------------------------//
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);


                xdoc = XDocument.Load(DefinitionsClass.CONFIG_FILE_PATH);
                if (xdoc == null)
                {

                    //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "-->{0}:{1} Failed to Read XMLConfigPath", GetType().Name, MethodBase.GetCurrentMethod().Name);
                }


                foreach (XElement element in xdoc.Descendants("CheckPrintReturn"))
                {
                    if (element.Value.ToUpper() == "TRUE")
                    {
                        g_CheckIsprintingStatus = true;
                        //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "-->{0}:{1} Printer Successful job check vlaue is -->  {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, "TRUE");

                    }
                    else if (element.Value.ToUpper() == "FALSE")
                    {
                        g_CheckIsprintingStatus = false;
                        //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "-->{0}:{1} Printer Successful job check vlaue is -->  {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, "FALSE");

                    }
                }
                /////////////////////////// check SuccessfulPrintCheckDuration
                foreach (XElement element in xdoc.Descendants("CheckPrintTimeout"))
                {
                    if ((element.Value == null) || (element.Value == ""))
                    {

                        //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "-->{0}:{1} SuccessfulPrintCheckDuration Not found in XML , Default value set  {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, m_IntSuccessfulPrintCheckDuration.ToString());

                    }
                    else if (Convert.ToInt32(element.Value) > 1000)
                    {
                        g_waitingUntillPriinterStartJob = Convert.ToInt32(element.Value);
                    }
                    else
                    {
                        //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "-->{0}:{1} SuccessfulPrintCheckDuration Value less than 1000 in XML , Default value set  {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, m_IntSuccessfulPrintCheckDuration.ToString());

                    }
                }


            }
            catch (Exception ex)
            {
                logger.Fatal(ex);
            }

        }

        public DefinitionsClass.PrinterReturnValue Printform(DefinitionsClass.PrintFormName printFormName, string[] strData)
        {

            //----------------------------LocalVar---------------------//
            // Represents an XML document. For the components and usage of an XDocument object
            XDocument formConfigXmlFile;
            int intCellCounter = 0;
            logger.Info("start Printform");
            string receiptFormFileAddress = string.Empty;
            string strFormat = string.Empty;
            DefinitionsClass.PrinterReturnValue printerReturnValue = DefinitionsClass.PrinterReturnValue.ERR_HARDWARE_ERROR;
            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                //find the form file address and returnthe address
                logger.Debug($"receipt form address:");
                receiptFormFileAddress = GetFormAddress(printFormName.ToString());
                logger.Debug($"receipt form address : {receiptFormFileAddress}");

                if (string.IsNullOrEmpty(receiptFormFileAddress))
                {
                    printerReturnValue = DefinitionsClass.PrinterReturnValue.ERR_INVALID_ADDRESS;
                    logger.Error("Form is empty, invalid address");
                    return printerReturnValue;
                }
                // set excell trial licence and save template form file address
                m_objExcelClass = new ExcelClass(receiptFormFileAddress);
                logger.Debug("Form excel created");
                //change xls to xml and return form config file address then load
                string xmlFormConfigFileAddress = GetFormXMLpath(receiptFormFileAddress);
                logger.Debug($"XMLCONFIGFILE {xmlFormConfigFileAddress}");
                if (string.IsNullOrEmpty(xmlFormConfigFileAddress))
                {
                    printerReturnValue = DefinitionsClass.PrinterReturnValue.ERR_INVALID_ADDRESS;
                    logger.Error("config filse is empty, invalid address");
                    return printerReturnValue;
                }
                formConfigXmlFile = XDocument.Load(xmlFormConfigFileAddress);
                logger.Debug($"form Config Xml File: {formConfigXmlFile}");
                foreach (XElement element in formConfigXmlFile.Descendants("cell"))
                {
                    intCellCounter++;
                }

                logger.Debug($"number of cells to print: {intCellCounter}");
                if (intCellCounter != strData.Length)
                {
                    printerReturnValue = DefinitionsClass.PrinterReturnValue.ERR_INTERNAL_ERROR;
                    logger.Error($"data not equal to cells: {intCellCounter}");
                    return printerReturnValue;
                }
                CellItems = new string[intCellCounter, 8];
                m_intLastRow = 0;
                foreach (XElement element in formConfigXmlFile.Descendants("cell"))
                {

                    if (element.Attribute("font") != null)
                    {
                        CellItems[m_intLastRow, 0] = element.Attribute("font").Value;
                        logger.Debug($"font: {CellItems[m_intLastRow, 0]}");
                    }
                    if (element.Attribute("index") != null)
                    {
                        CellItems[m_intLastRow, 1] = element.Attribute("index").Value;
                        logger.Debug($"index: {CellItems[m_intLastRow, 1]}");
                    }
                    if (element.Attribute("bold") != null)
                    {
                        CellItems[m_intLastRow, 2] = element.Attribute("bold").Value;
                        logger.Debug($"bold: {CellItems[m_intLastRow, 2]}");
                    }
                    if (element.Attribute("size") != null)
                    {
                        CellItems[m_intLastRow, 3] = element.Attribute("size").Value;
                        logger.Debug($"size: {CellItems[m_intLastRow, 3]}");
                    }


                    CellItems[m_intLastRow, 4] = strData[m_intLastRow];    // Text
                    logger.Debug($"text: {CellItems[m_intLastRow, 4]}");


                    if (element.Attribute("FORMAT") != null)
                    {
                        CellItems[m_intLastRow, 5] = element.Attribute("FORMAT").Value;
                        logger.Debug($"FORMAT: {CellItems[m_intLastRow, 5]}");
                        CellItems[m_intLastRow, 4] = CellItems[m_intLastRow, 4].Replace(" ", "")
                                        .Replace("\t", "").Replace("\n", "").Replace("\r", "");

                        logger.Debug($"FORMATTED Text: {CellItems[m_intLastRow, 4]}");
                        logger.Debug($"To be Formatted: {CellItems[m_intLastRow, 5]}");
                        switch (CellItems[m_intLastRow, 5])
                        {
                            case "Time":
                                {
                                    var strStringBuilder = new StringBuilder(CellItems[m_intLastRow, 4]);
                                    logger.Debug($"time before: {strStringBuilder}");
                                    strStringBuilder.Insert(2, ":");//(18:3654)
                                    if (strStringBuilder.Length>5)
                                    {
                                        strStringBuilder.Insert(5, ":");//(18:36:54)
                                    }
                                    CellItems[m_intLastRow, 4] = strStringBuilder.ToString();
                                    logger.Debug($"time after: {strStringBuilder}");
                                }
                                break;
                            case "Date":
                                {

                                    var strStringBuilder = new StringBuilder(CellItems[m_intLastRow, 4]);
                                    logger.Debug($"Date before: {strStringBuilder}");
                                    if (CellItems[m_intLastRow, 4].Length == 8)
                                    {
                                        strStringBuilder.Insert(4, "/");//(1401/0105)
                                        strStringBuilder.Insert(7, "/");//(1401/01/05)
                                    }
                                    if (CellItems[m_intLastRow, 4].Length == 4)
                                    {
                                        strStringBuilder.Insert(2, "/");//(01/01)
                                    }
                                    if (CellItems[m_intLastRow, 4].Length == 6)
                                    {
                                        strStringBuilder.Insert(2, "/");//(01/0105)
                                        strStringBuilder.Insert(5, "/");//(01/01/05)
                                    }

                                    CellItems[m_intLastRow, 4] = strStringBuilder.ToString();
                                    logger.Debug($"Date after: {strStringBuilder}");

                                }
                                break;
                            case "Pan":
                                {
                                    var strStringBuilder = new StringBuilder(CellItems[m_intLastRow, 4]);
                                    logger.Debug($"Pan before: {strStringBuilder}");
                                    try
                                    {
                                        strStringBuilder.Remove(6, 6);
                                        strStringBuilder.Insert(6, "******");
                                    CellItems[m_intLastRow, 4] = strStringBuilder.ToString();
                                        logger.Debug($"Pan after: {strStringBuilder}");
                                    }
                                    catch (Exception)
                                    {
                                        CellItems[m_intLastRow, 4] = CellItems[m_intLastRow, 4];
                                    }
                                }
                                break;
                            case "Currency":
                                {
                                    StringBuilder strStringBuilder = new StringBuilder(CellItems[m_intLastRow, 4]);
                                    logger.Debug($"currency before: {strStringBuilder}");
                                    int Length = CellItems[m_intLastRow, 4].Length;
                                    int symCounter = Length / 3;
                                    // 100 000 000 -> Length 9 -> symCounter 3
                                    for (int i = 1; i <= symCounter; i++)
                                    {
                                        int iSymPosition = Length - (i * 3);
                                        if (iSymPosition>0)
                                        {
                                            strStringBuilder.Insert(iSymPosition, ",");
                                        }
                                    }
                                    CellItems[m_intLastRow, 4] = strStringBuilder.ToString();
                                    logger.Debug($"time after: {strStringBuilder}");
                                }
                                break;
                        }
                    }
                    m_intLastRow++;
                    logger.Debug($"row number: {m_intLastRow}");
                }
                logger.Debug($"open excel file: ");
                m_objExcelClass.OpenExcelFile();
                logger.Debug($"write data to cells: ");
                m_objExcelClass.WriteToCell(CellItems);

                logger.Debug($"set print configures: ");
                m_objExcelClass.SetPrinterConfig(xmlFormConfigFileAddress);

                logger.Debug($"print form file: ");
                m_objExcelClass.PrintMyExcelFile(receiptFormFileAddress);

                m_intLastRow = 0;


                if (g_SuccessfulPrint)
                {
                    printerReturnValue = DefinitionsClass.PrinterReturnValue.SUCCESS;
                    //LOG HERE

                    logger.Debug($"Printe success: {g_SuccessfulPrint}");
                }
                else
                {
                    printerReturnValue = DefinitionsClass.PrinterReturnValue.ERR_DEV_NOT_READY;
                    //LOG HERE

                    logger.Error($" in value return : {g_SuccessfulPrint}");
                }


            }
            catch (Exception ex)
            {

                //m_objTrace.Error(TraceLog.ERR_LEVEL.CRITICAL, "<->{0}:{1} failed.\n\r Error: {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, ex);
                logger.Error(ex.Message);
                logger.Fatal(ex);
                printerReturnValue = DefinitionsClass.PrinterReturnValue.ERR_FRM_FATAL_ERROR;
            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} Return Value is {2}.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name, objRet);
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);
            logger.Debug($"return {printerReturnValue}");
            return printerReturnValue;
        }

        private string GetFormAddress(string strFormName)
        {
            //----------------------------LocalVar---------------------//
            XDocument formsFileAddressList = null;
            string Retval = string.Empty;
            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                //m_objTrace.Trace(TraceLog.TRC_LEVEL.PARAM, "<->{0}:{1} input Params is :FormNo:{2}", GetType().Name, MethodBase.GetCurrentMethod().Name, strFormNo);
                formsFileAddressList = XDocument.Load(m_strXMLConfigPath);
                foreach (XElement element in formsFileAddressList.Descendants("Folder"))
                {
                    Retval = element.Attribute(DefinitionsClass.STR_FORM_PATH).Value;
                }
                foreach (XElement element in formsFileAddressList.Descendants("Form"))
                {
                    if (element.Attribute(DefinitionsClass.STR_FORM_NAME).Value == strFormName)
                    {
                        Retval += element.Attribute(DefinitionsClass.STR_FORM_FILE).Value;
                        //m_objTrace.Trace(TraceLog.TRC_LEVEL.DEBUG, "<->{0}:{1} Find Form !\r\nRetval:{2}", GetType().Name, MethodBase.GetCurrentMethod().Name, Retval);
                        break;
                    }
                }

            }
            catch (Exception ex)
            {
                Retval = "";
                logger.Fatal(ex.Message);
                logger.Fatal(ex);
            }

            logger.Debug($"return value {Retval}");
            return Retval;
        }
        private string GetFormXMLpath(string strPath)
        {
            //----------------------------LocalVar---------------------//
            string strRetVal = "";
            //-------------------------------------------------------//
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "-->{0}:{1} Call", GetType().Name, MethodBase.GetCurrentMethod().Name);
            try
            {
                strRetVal = strPath.Replace(".xlsx", ".xml");
            }
            catch (Exception ex)
            {
                logger.Fatal(ex.Message);
                logger.Fatal(ex);
                //m_objTrace.Error(TraceLog.ERR_LEVEL.CRITICAL, "<->{0}:{1} failed.\n\r Error: {2}", GetType().Name, MethodBase.GetCurrentMethod().Name, ex);

            }
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} Return Value is {2}.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name, strRetVal);
            //m_objTrace.Trace(TraceLog.TRC_LEVEL.CALL, "<-{0}:{1} End.\n\r", GetType().Name, MethodBase.GetCurrentMethod().Name);

            logger.Debug($"return value {strRetVal}"); 
            return strRetVal;
        }

    }
}
