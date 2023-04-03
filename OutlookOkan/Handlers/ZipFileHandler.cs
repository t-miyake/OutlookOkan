using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;

namespace OutlookOkan.Handlers
{
    public sealed class ZipFileHandler
    {
        internal readonly List<string> IncludeExtensions = new List<string>();

        /// <summary>
        /// 暗号化ZIPファイル(パスワード付きZIP)か否かを判定する。
        /// </summary>
        /// <param name="filePath">確認したいファイルのフルパス</param>
        /// <returns>暗号化ZIPか否か</returns>
        internal bool CheckZipIsEncryptedAndGetIncludeExtensions(string filePath)
        {
            //リンクとして添付の場合、実ファイルが存在しない場合がある。
            if (!File.Exists(filePath)) return false;

            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var zipInputStream = new ZipInputStream(fileStream);

            var isEncrypted = false;

            while (true)
            {
                ZipEntry zip;
                try
                {
                    zip = zipInputStream.GetNextEntry();
                    try
                    {
                        if (zip != null)
                        {
                            IncludeExtensions.Add(Path.GetExtension(zip.Name.ToLower()));
                        }
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }

                }
                catch (ZipException)
                {
                    break;
                }
                catch (NotSupportedException)
                {
                    isEncrypted = true;
                    break;
                }
                catch (NullReferenceException)
                {
                    break;
                }
                catch (InvalidOperationException)
                {
                    break;
                }

                if (zip?.IsCrypted ?? false)
                {
                    isEncrypted = true;
                    break;
                }
            }

            zipInputStream.Dispose();

            return isEncrypted;
        }
    }
}