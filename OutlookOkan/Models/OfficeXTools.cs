using Microsoft.Office.Core;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookOkan.Models
{
    public sealed class OfficeXTools
    {
        internal bool CheckOfficeXIsEncrypted(string filePath, string fileType)
        {
            //リンクとして添付の場合、実ファイルが存在しない場合がある。
            if (!File.Exists(filePath)) return false;
            var isEncrypted = false;

            var thisFileType = fileType.ToLower().Replace(".", "");
            switch (thisFileType)
            {
                case "doc":
                case "docx":
                    var tempWordApp = new Word.Application
                    {
                        Application =
                        {
                            AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                        },
                        Visible = false
                    };

                    try
                    {
                        var wordFile = tempWordApp.Documents.Open(filePath, PasswordDocument: "unknown", Visible: false);
                        isEncrypted = false;

                        Thread.Sleep(10);
                        wordFile.Close();
                        Thread.Sleep(10);
                        _ = Marshal.ReleaseComObject(wordFile);
                        wordFile = null;
                    }
                    catch (Exception e)
                    {
                        //パスワード違いの例外となった場合、パスワード付きDOCXとして判定。
                        isEncrypted = e.HResult == -2146822880;
                    }

                    Thread.Sleep(10);
                    tempWordApp.Quit();
                    Thread.Sleep(10);
                    _ = Marshal.ReleaseComObject(tempWordApp);
                    tempWordApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    break;

                case "xls":
                case "xlsx":
                    var tempExcelApp = new Excel.Application
                    {
                        Application =
                        {
                            EnableEvents = false,
                            AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable
                        },
                        Visible = false
                    };

                    try
                    {
                        var excelFile = tempExcelApp.Workbooks.Open(filePath, Password: null);
                        isEncrypted = false;
                        Thread.Sleep(10);
                        excelFile.Close(false);
                        Thread.Sleep(10);
                        _ = Marshal.ReleaseComObject(excelFile);
                        excelFile = null;
                    }
                    catch (Exception e)
                    {
                        //パスワード違いの例外となった場合、パスワード付きXLSXとして判定。
                        isEncrypted = e.HResult == -2146827284;
                    }

                    Thread.Sleep(10);
                    tempExcelApp.Quit();
                    Thread.Sleep(10);
                    _ = Marshal.ReleaseComObject(tempExcelApp);
                    tempExcelApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    break;

                case "ppt":
                case "pptx":
                    var tempPowerPointApp = new PowerPoint.Application();

                    try
                    {
                        var pptFile = tempPowerPointApp.Presentations.Open(filePath + "::unknown", MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                        isEncrypted = false;

                        Thread.Sleep(10);
                        pptFile.Close();
                        Thread.Sleep(10);
                        _ = Marshal.ReleaseComObject(pptFile);
                        pptFile = null;
                    }
                    catch (Exception e)
                    {
                        //パスワード違いの例外となった場合、パスワード付きPPTXとして判定。
                        isEncrypted = e.HResult == -2147467259;
                    }

                    Thread.Sleep(10);
                    tempPowerPointApp.Quit();
                    Thread.Sleep(10);
                    _ = Marshal.ReleaseComObject(tempPowerPointApp);
                    tempPowerPointApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    break;

                default:
                    return false;
            }

            return isEncrypted;
        }
    }
}