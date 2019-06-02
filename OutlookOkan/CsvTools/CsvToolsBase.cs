using CsvHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookOkan.CsvTools
{
    public class CsvToolsBase
    {
        /// <summary>
        /// 文字コードの確認
        /// </summary>
        /// <param name="filePath">文字コードを確認するファイルのパス</param>
        /// <returns>文字コード</returns>
        internal string DetectCharset(string filePath)
        {
            try
            {
                using (var fileStream = File.OpenRead(filePath))
                {
                    var charsetDetector = new Ude.CharsetDetector();
                    charsetDetector.Feed(fileStream);
                    charsetDetector.DataEnd();

                    return charsetDetector.Charset ?? "UTF-8";
                }
            }
            catch (Exception)
            {
                return "UTF-8";
            }
        }

        /// <summary>
        /// 読み込んだCSVから、List<T/>を返す。
        /// </summary>
        /// <typeparam name="TCsvType"></typeparam>
        /// <param name="loadedCsv">読み込んだCSVデータ</param>
        /// <returns>CSVデータ(List<T/>)</returns>
        public List<TCsvType> GetCsvRecords<TCsvType>(CsvReader loadedCsv)
        {
            loadedCsv.Configuration.MissingFieldFound = null;
            var list = loadedCsv.GetRecords<TCsvType>().ToList();
            loadedCsv.Dispose();

            return list;
        }
    }
}