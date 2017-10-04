using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CsvHelper;

namespace OutlookOkan
{
    public class ReadAndWriteCsv
    {
        /// <summary>
        /// 設定ファイル(CSV)の設置個所は下記で固定。
        /// C:\Users\USERNAME\AppData\Roaming\Noraneko\OutlookOkan\
        /// </summary>
        private readonly string _directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
        private readonly string _filePath;

        public ReadAndWriteCsv(string filename)
        {
            _filePath = _directoryPath + (filename ?? throw new ArgumentNullException(nameof(filename)));

            CheckFileAndDirectoryExists();
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
        /// CSVファイルを読み込みパースする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <returns>パースされたCSV</returns>
        public CsvParser ParseCsv<TMaptype>() where TMaptype : CsvHelper.Configuration.CsvClassMap
        {
            var csvParser = new CsvParser(new StreamReader(_filePath, Encoding.GetEncoding("Shift_JIS")));
            csvParser.Configuration.HasHeaderRecord = false;
            csvParser.Configuration.RegisterClassMap<TMaptype>();

            return csvParser;
        }

        /// <summary>
        /// パースされたCSVをもとに、List<T/>を変えす。
        /// </summary>
        /// <typeparam name="TCsvType"></typeparam>
        /// <param name="csvPerser">パースされたCSV</param>
        /// <returns>CSVデータ(List<T/>)</returns>
        public List<TCsvType> ReadCsv<TCsvType>(CsvParser csvPerser)
        { 
            var list = new CsvReader(csvPerser).GetRecords<TCsvType>().ToList();
            csvPerser.Dispose();

            return list;
        }

        /// <summary>
        /// BindingSource型のデータをCSVファイルに書き込む。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="bindableData">BindingSource型のデータ</param>
        public void WriteBindableDataToCsv<TMaptype>(BindingSource bindableData) where TMaptype : CsvHelper.Configuration.CsvClassMap
        {
            var csvWriter = new CsvWriter(new StreamWriter(_filePath, false, Encoding.GetEncoding("Shift_JIS")));
            csvWriter.Configuration.HasHeaderRecord = false;
            csvWriter.Configuration.RegisterClassMap<TMaptype>();

            csvWriter.WriteRecords(bindableData);

            csvWriter.Dispose();
        }
    }
}