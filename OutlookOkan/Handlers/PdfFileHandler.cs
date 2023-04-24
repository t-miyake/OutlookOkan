using PdfSharp.Pdf.IO;
using System;
using System.IO;

namespace OutlookOkan.Handlers
{
    internal static class PdfFileHandler
    {
        internal static bool CheckPdfIsEncrypted(string filePath)
        {
            //リンクとして添付の場合、実ファイルが存在しない場合がある。
            if (!File.Exists(filePath)) return false;

            try
            {
                PdfReader.Open(filePath, PdfDocumentOpenMode.ReadOnly).Dispose();
            }
            catch (PdfReaderException)
            {
                return true;
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }
    }
}