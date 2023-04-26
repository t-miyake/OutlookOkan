using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Win32;
using OutlookOkan.Properties;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace OutlookOkan.Handlers
{
    internal static class CsvFileHandler
    {
        private static readonly CsvConfiguration Config = new CsvConfiguration(CultureInfo.CurrentCulture)
        {
            HasHeaderRecord = false,
            MissingFieldFound = null
        };

        /// <summary>
        /// 設定ファイル(CSV)の設置個所は下記で固定。
        /// C:\Users\USERNAME\AppData\Roaming\Noraneko\OutlookOkan\
        /// </summary>
        private static readonly string DirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");

        internal static List<T> ReadCsv<T>(Type classMapType, string fileName) where T : class
        {
            var fullPath = Path.Combine(DirectoryPath, fileName);
            if (!File.Exists(fullPath))
            {
                return new List<T>();
            }

            using (var reader = new StreamReader(fullPath, Encoding.UTF8))
            using (var csv = new CsvReader(reader, Config))
            {
                csv.Context.RegisterClassMap(classMapType);
                return csv.GetRecords<T>().ToList();
            }
        }

        internal static void AppendCsv(Type classMapType, string fileName, object record)
        {
            var fullPath = Path.Combine(DirectoryPath, fileName);

            if (!File.Exists(fullPath))
            {
                var records = new List<object> { record };
                CreateOrReplaceCsv(classMapType, fileName, records);
                return;
            }

            using (var writer = new StreamWriter(fullPath, true, Encoding.UTF8))
            using (var csv = new CsvWriter(writer, Config))
            {
                csv.Context.RegisterClassMap(classMapType);
                csv.WriteRecord(record);
                writer.Flush();
            }
        }

        internal static void CreateOrReplaceCsv(Type classMapType, string fileName, IEnumerable<object> records, string directoryPath = null)
        {
            var targetDirectory = directoryPath ?? DirectoryPath;
            if (!Directory.Exists(targetDirectory))
            {
                Directory.CreateDirectory(targetDirectory);
            }

            var fullPath = Path.Combine(targetDirectory, fileName);

            using (var writer = new StreamWriter(fullPath, false, Encoding.UTF8))
            using (var csv = new CsvWriter(writer, Config))
            {
                csv.Context.RegisterClassMap(classMapType);
                csv.WriteRecords(records);
                writer.Flush();
            }
        }

        internal static List<T> ImportCsv<T>(Type classMapType) where T : class
        {
            _ = MessageBox.Show(Resources.CSVImportAlert, Resources.AppName, MessageBoxButton.OK);

            var openFileDialog = new OpenFileDialog
            {
                Title = Resources.SelectCSVFile,
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv"
            };

            var importPath = openFileDialog.ShowDialog() ?? false ? openFileDialog.FileName : null;
            if (importPath is null) return new List<T>();

            try
            {

                using (var reader = new StreamReader(importPath, Encoding.UTF8))
                using (var csv = new CsvReader(reader, Config))
                {
                    csv.Context.RegisterClassMap(classMapType);
                    return csv.GetRecords<T>().ToList();
                }
            }
            catch (CsvHelperException ex)
            {
                _ = MessageBox.Show(Resources.ImportFailed + ex, Resources.AppName, MessageBoxButton.OK);
                return new List<T>();
            }
        }

        internal static void ExportCsv(Type classMapType, IEnumerable<object> records, string defaultFileName)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = Resources.SelectSaveDestination,
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv",
                CreatePrompt = true,
                FileName = defaultFileName
            };

            if (!saveFileDialog.ShowDialog() ?? false) return;
            try
            {
                CreateOrReplaceCsv(classMapType, saveFileDialog.FileName, records);
                _ = MessageBox.Show(Resources.SuccessfulExport, Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception e)
            {
                _ = MessageBox.Show(Resources.ExportFailed + e, Resources.AppName, MessageBoxButton.OK);
            }
        }
    }
}