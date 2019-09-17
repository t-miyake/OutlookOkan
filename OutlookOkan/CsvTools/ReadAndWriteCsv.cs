using CsvHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OutlookOkan.CsvTools
{
    public sealed class ReadAndWriteCsv : CsvToolsBase
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

            CheckFileAndDirectoryExists();

            _fileEncoding = Encoding.GetEncoding(DetectCharset(_filePath));
        }

        /// <summary>
        /// 指定パスにおけるファイルやフォルダの有無を確認し、無い場合はそれぞれ作成する。
        /// </summary>
        private void CheckFileAndDirectoryExists()
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
        /// データをCSVファイルに書き込む。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="records">ArrayList型のデータ</param>
        public void WriteRecordsToCsv<TMaptype>(List<object> records) where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var csvWriter = new CsvWriter(new StreamWriter(_filePath, false, Encoding.UTF8));
            csvWriter.Configuration.HasHeaderRecord = false;
            csvWriter.Configuration.RegisterClassMap<TMaptype>();

            csvWriter.WriteRecords(records);

            csvWriter.Dispose();
        }
    }
}