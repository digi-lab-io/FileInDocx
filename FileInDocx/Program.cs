/*
Copyright 2015 Softcom
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
   http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

**/

using System;
using System.Reflection;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Microsoft.Win32;

namespace FileInDocx
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting object embedding...");

            //Set the Word Application Window Title
            string wordAppId = "" + DateTime.Now.Ticks;

            Word.Application word = new Word.Application();
            word.Application.Caption = wordAppId;
            word.Application.Visible = true;
            int processId = GetProcessIdByWindowTitle(wordAppId);
            word.Application.Visible = false;

            try
            {
                object missing = Missing.Value;
                string IconIndex = "2";
                string IconFilePath = @"%SystemRoot%\system32\imageres.dll";
                try
                {
                    //opens the registry for the wanted key.
                    RegistryKey Root = Registry.ClassesRoot;
                    if (Root != null)
                    {
                        string ext = Path.GetExtension(args[0]);
                        if (!string.IsNullOrEmpty(ext))
                        {
                            RegistryKey ExtensionKey = Root.OpenSubKey(ext);
                            if (ExtensionKey != null)
                            {
                                ExtensionKey.GetValueNames();
                                RegistryKey ApplicationKey = Root.OpenSubKey(ExtensionKey.GetValue("").ToString());

                                if (ApplicationKey != null)
                                {
                                    //gets the name of the file that have the icon.
                                    string IconLocation = ApplicationKey.OpenSubKey("DefaultIcon").GetValue("").ToString();
                                    string[] IconPath = IconLocation.Split(',');

                                    if (File.Exists(IconPath[0]))
                                    {
                                        IconFilePath = IconPath[0];
                                        if (IconPath.Length > 1 && IconPath[1] != null)
                                        {
                                            IconIndex = IconPath[1];
                                        }
                                        else
                                        {
                                            IconIndex = "0";
                                        }
                                        Console.WriteLine("Icon found: {0}, {1}", IconPath[0], IconIndex);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error reading registry: {0}", e);
                }

                Word.Document document = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                document.InlineShapes.AddOLEObject(ref missing, args[0], false, true, IconFilePath, Convert.ToInt16(IconIndex), Path.GetFileName(args[0]), ref missing);
                document.SaveAs(args[1], Word.WdSaveFormat.wdFormatDocumentDefault);
                document.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error ocurred: {0}", e);
            }
            finally
            {
                // Terminate Winword instance by PID.
                Console.WriteLine("Terminating Winword process with the Windowtitle '{0}' and the Application ID: '{1}'.", wordAppId, processId);
                Process process = Process.GetProcessById(processId);
                process.Kill();
            }
        }

        public static int GetProcessIdByWindowTitle(string paramWordAppId)
        {
            Process[] P_CESSES = Process.GetProcessesByName("WINWORD");
            for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
            {
                if (P_CESSES[p_count].MainWindowTitle.Equals(paramWordAppId))
                {
                    return P_CESSES[p_count].Id;
                }
            }
            return Int32.MaxValue;
        }
    }
}