using CsvHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookOkan.CsvTools
{
    public class ReadAndWriteCsv
    {
        /// <summary>
        /// 設定ファイル(CSV)の設置個所は下記で固定。
        /// C:\Users\USERNAME\AppData\Roaming\Noraneko\OutlookOkan\
        /// </summary>
        private readonly string _directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
        private readonly string _filePath;
        private readonly Encoding _fileEncoding;

        public ReadAndWriteCsv(string filename)
        {
            _filePath = _directoryPath + (filename ?? throw new ArgumentNullException(nameof(filename)));
            _fileEncoding = Encoding.GetEncoding(DetectCharset(_filePath));

            CheckFileAndDirectoryExists();
        }

        /// <summary>
        /// 文字コードの確認
        /// </summary>
        /// <param name="filePath">文字コードを確認するファイルのパス</param>
        /// <returns>文字コード</returns>
        public string DetectCharset(string filePath)
        {
            using (var fileStream = File.OpenRead(filePath))
            {
                var charsetDetector = new Ude.CharsetDetector();
                charsetDetector.Feed(fileStream);
                charsetDetector.DataEnd();

                return charsetDetector.Charset ?? "UTF-8";
            }
        }

        /// <summary>
        /// 指定パスにおけるファイルやフォルダの有無を確認し、無い場合はそれぞれ作成する。
        /// </summary>
        public void CheckFileAndDirectoryExists()
        {
            if (!Directory.Exists(_directoryPath))
                Directory.CreateDirectory(_directoryPath);

            if (!File.Exists(_filePath))
                File.Create(_filePath).Close();
        }

        /// <summary>
        /// CSVファイルを読み込む
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <returns>読み込んだCSVデータ</returns>
        public CsvReader LoadCsv<TMaptype>() where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var csvReader = new CsvReader(new StreamReader(_filePath, _fileEncoding));
            csvReader.Configuration.HasHeaderRecord = false;
            csvReader.Configuration.RegisterClassMap<TMaptype>();

            return csvReader;
        }

        /// <summary>
        /// 読み込んだCSVから、List<T/>を変えす。
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

        /// <summary>
        /// BindingSource型のデータをCSVファイルに書き込む。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="bindableData">BindingSource型のデータ</param>
        public void WriteBindableDataToCsv<TMaptype>(BindingSource bindableData) where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var csvWriter = new CsvWriter(new StreamWriter(_filePath, false, Encoding.UTF8));
            csvWriter.Configuration.HasHeaderRecord = false;
            csvWriter.Configuration.RegisterClassMap<TMaptype>();

            csvWriter.WriteRecords(bindableData);

            csvWriter.Dispose();
        }
    }
}