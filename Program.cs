using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using WinShell;

namespace Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;

                var test = new TestPropertyStore();
                test.PerformTests();
            }
            catch (Exception err)
            {
#if DEBUG
                Console.WriteLine(err.ToString());
#else
                Console.WriteLine(err.Message);
#endif
            }

            Win32Interop.ConsoleHelper.PromptAndWaitIfSoleConsole();
        }
    }

    class TestPropertyStore
    {
        const string c_SampleJpg = "sample.jpg";
        string m_workingDirectory;

        public TestPropertyStore()
        {
            // Find the working directory which is where sample.jpg is located.
            // It should be the current directory or a parent thereof.

            string workingDirectory = Environment.CurrentDirectory;
            while (!File.Exists(Path.Combine(workingDirectory, c_SampleJpg)))
            {
                workingDirectory = Path.GetDirectoryName(workingDirectory);
                if (string.IsNullOrEmpty(workingDirectory))
                {
                    throw new ApplicationException(string.Format("Test file, '{0}', not found in working directory path.", c_SampleJpg));
                }
                m_workingDirectory = workingDirectory;
            }
        }

        public void PerformTests()
        {
            Console.WriteLine("Working directory: " + m_workingDirectory);
            Console.WriteLine();

            /*
            DumpAllProperties(Path.Combine(m_workingDirectory, "src.mp3"));
            Console.WriteLine();
            DumpAllProperties(Path.Combine(m_workingDirectory, "src.jpg"));
            Console.WriteLine();

            PerformCopyTest(".jpg");
            Console.WriteLine();
            PerformCopyTest(".mp3");
            Console.WriteLine();
            */

            try
            {
                m_ps = new PropertySystem();
                m_log = new StreamWriter(Path.Combine(m_workingDirectory, "testLog.txt"), false);
                m_foundKeys = new HashSet<PROPERTYKEY>();
                m_errorKeys = new HashSet<PROPERTYKEY>();

                RetrieveAll(@"C:\");
            }
            finally
            {
                m_foundKeys = null;
                if (m_log != null)
                {
                    m_log.Dispose();
                    m_log = null;
                }
                if (m_ps != null)
                {
                    m_ps.Dispose();
                    m_ps = null;
                }
            }

        }

        void PerformCopyTest(string ext)
        {
            string srcFilename = Path.Combine(m_workingDirectory, string.Concat("src", ext));
            string blankFilename = Path.Combine(m_workingDirectory, string.Concat("blank", ext));
            string dstFilename = Path.Combine(m_workingDirectory, string.Concat("dst", ext));

            File.Copy(blankFilename, dstFilename, true);
            CopyAllProperties(srcFilename, dstFilename);
        }

        PropertySystem m_ps = null;
        TextWriter m_log = null;
        HashSet<PROPERTYKEY> m_foundKeys = null;
        HashSet<PROPERTYKEY> m_errorKeys = null;

        public void RetrieveAll(string path)
        {
            try
            {
                Console.WriteLine(path);
                foreach (string filePath in Directory.GetFiles(path))
                {
                    PropertyStore propStore = null;
                    try
                    {
                        bool hasError = false;
                        try
                        {
                            propStore = PropertyStore.Open(filePath);
                        }
                        catch (Exception err)
                        {
                            Console.WriteLine("Failed to open property store on: {0}\r\n{1}\r\n", filePath, err.ToString());
                            // m_log.WriteLine("Failed to open property store on: {0}\r\n{1}\r\n", filePath, err.ToString());
                            hasError = true;
                        }

                        if (!hasError)
                        {
                            int count = propStore.Count;
                            for (int i = 0; i < propStore.Count; ++i)
                            {
                                PROPERTYKEY propKey;
                                try
                                {
                                    // Get the key for the enumerated property
                                    propKey = propStore.GetAt(i);
                                }
                                catch (Exception err)
                                {
                                    Console.WriteLine("Failed to retrieve property key on '{0}' index='{1}'\r\n{2}\r\n", filePath, i, err.ToString());
                                    m_log.WriteLine("Failed to retrieve property key on '{0}' index='{1}'\r\n{2}\r\n", filePath, i, err.ToString());
                                    continue;
                                }

                                try
                                {
                                    // Get the value
                                    object value = propStore.GetValue(propKey);
                                }
                                catch (Exception err)
                                {
                                    if (m_errorKeys.Add(propKey))
                                    {
                                        string message = (err.InnerException != null) ? err.InnerException.Message : err.Message;

                                        // Attempt to get the canonical name
                                        string name = string.Empty;
                                        try
                                        {
                                            name = m_ps.GetPropertyDescription(propKey).CanonicalName;
                                        }
                                        catch
                                        {
                                            name = string.Empty;
                                        }

                                        Console.WriteLine("Failed to retrieve value. file='{0}' propkey='{1}' canonicalName='{2}'\r\n{3}\r\n", filePath, propKey, name, message);
                                        m_log.WriteLine("Failed to retrieve value. file='{0}' propkey='{1}' canonicalName='{2}'\r\n{3}\r\n", filePath, propKey, name, message);
                                    }
                                }

                                if (m_foundKeys.Add(propKey))
                                {
                                    PropertyDescription desc = null;
                                    try
                                    {
                                        // Get the description from the property store (if available)
                                        desc = m_ps.GetPropertyDescription(propKey);
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine("Error retrieving property description. propkey='{0}'\r\n{1}\r\n", propKey, err.ToString());
                                        m_log.WriteLine("Error retrieving property description. propkey='{0}'\r\n{1}\r\n", propKey, err.ToString());
                                    }
                                    if (desc != null)
                                    {
                                        if (string.IsNullOrEmpty(desc.CanonicalName))
                                        {
                                            Console.WriteLine("No canonical name provided. propkey='{0}'\r\n", propKey);
                                            m_log.WriteLine("No canonical name provided. propkey='{0}'\r\n", propKey);
                                        }
                                        else if (string.IsNullOrEmpty(desc.DisplayName))
                                        {
                                            Console.WriteLine("No display name provided. propkey='{0}' canonical='{1}'\r\n", propKey, desc.CanonicalName);
                                            m_log.WriteLine("No display name provided. propkey='{0}' canonical='{1}'\r\n", propKey, desc.CanonicalName);
                                        }
                                    }
                                }

                            }
                        }
                    }
                    finally
                    {
                        if (propStore != null)
                        {
                            propStore.Dispose();
                            propStore = null;
                        }
                    }
                } // foreach file

                // Recursively enumerate directories
                foreach (string dirPath in Directory.GetDirectories(path))
                {
                    RetrieveAll(dirPath);
                }
            }
            catch (UnauthorizedAccessException err)
            {
                Console.WriteLine("Error: " + err.ToString() + "\r\n");
            }
            catch (Exception err)
            {
                Console.WriteLine("Error: " + err.ToString() + "\r\n");
                m_log.WriteLine("Error: " + err.ToString() + "\r\n");
            }

        }

        /// <summary>
        /// Sample code showing how to dump all properties on a file including looking up
        /// the names of the properties.
        /// </summary>
        /// <param name="filename">Name of the file on which to dump all properties.</param>
        public static void DumpAllProperties(string filename)
        {
            Console.WriteLine("All properties in '{0}'.", filename);

            var propList = new List< KeyValuePair<string, string> >();
            using (var propSys = new PropertySystem())
            {
                using (var propStore = PropertyStore.Open(filename))
                {
                    int count = propStore.Count;
                    for (int i=0; i< count; ++i)
                    {
                        // Get the key for the enumerated property
                        PROPERTYKEY propKey = propStore.GetAt(i);

                        // Get the description from the property store (if available)
                        var desc = propSys.GetPropertyDescription(propKey);

                        // Get the value
                        object value = null;
                        try
                        {
                            value = propStore.GetValue(propKey);
                        }
                        catch (NotImplementedException err)
                        {
                            value = err.Message;
                        }

                        string valueStr = ValueToString(value);

                        // Retrieve or generate a name
                        string name;
                        if (desc != null)
                        {
                            name = string.Format("{0} ({1}) ({2}, {3})", desc.CanonicalName, desc.DisplayName,
                                (ShortPropTypeFlags)desc.TypeFlags, propStore.IsPropertyWriteable(propKey) ? "Writable" : "ReadOnly");
                        }
                        else
                        {
                            name = string.Format("{0}, {1}", propKey.fmtid, propKey.pid);
                        }

                        // Add to the list
                        propList.Add(new KeyValuePair<string, string>(name, valueStr));
                    }
                }
            }

            // Sort the list by key
            propList.Sort((a, b) => a.Key.CompareTo(b.Key) );

            // Dump the list
            foreach(var pair in propList)
            {
                Console.WriteLine("{0}: {1}", pair.Key, pair.Value);
            }

        } // DumpAllProperties

        private static string ValueToString(object value)
        {
            if (value == null) return "(null)";

            Array arr = value as Array;
            if (arr != null)
            {
                StringBuilder builder = new StringBuilder();
                foreach (var val in arr)
                {
                    if (builder.Length > 0) builder.Append(", ");
                    builder.Append((val != null) ? val.ToString() : "(null)");
                }
                return builder.ToString();
            }

            if (value is DateTime)
            {
                return ((DateTime)value).ToString("O", System.Globalization.CultureInfo.InvariantCulture);
            }

            return value.ToString();
        }

        public static void CopyAllProperties(string srcFilename, string dstFilename)
        {
            Console.WriteLine("Copying properties from '{0}' to '{1}'.", srcFilename, dstFilename);

            using (var propSys = new PropertySystem())
            {
                using (var srcPs = PropertyStore.Open(srcFilename))
                {
                    using (var dstPs = PropertyStore.Open(dstFilename, true))
                    {
                        int count = srcPs.Count;
                        for (int i = 0; i < count; ++i)
                        {
                            // Get the key for the enumerated property
                            PROPERTYKEY propKey = srcPs.GetAt(i);

                            // Get the description from the property store (if available)
                            var desc = propSys.GetPropertyDescription(propKey);
                            if (desc == null) continue; // Don't copy properties without descriptions

                            // Only copy if the property is not innate
                            if ((desc.TypeFlags & PROPDESC_TYPE_FLAGS.PDTF_ISINNATE) != 0)
                            {
                                Console.WriteLine("Property '{0}' is innate.", desc.CanonicalName);
                            }
                            else if (!dstPs.IsPropertyWriteable(propKey))
                            {
                                Console.WriteLine("Property '{0}' is read-only.", desc.CanonicalName);
                            }
                            else
                            {
                                Console.WriteLine("Copying property '{0}' ({1}).", desc.CanonicalName, desc.TypeFlags);

                                // Read the value
                                object value;
                                try
                                {
                                    value = srcPs.GetValue(propKey);
                                }
                                catch (NotImplementedException err) // Let all other exceptiosn go through
                                {
                                    Console.WriteLine(err.Message);
                                    continue;
                                }

                                // Write the value
                                try
                                {
                                    dstPs.SetValue(propKey, value);
                                }
                                catch
                                {
                                    Console.WriteLine("PropertyKey: {0}", propKey);
                                    throw;
                                }
                            }
                        }

                        Console.WriteLine("Committing changes.");
                        dstPs.Commit();
                    }
                }
            }

        }

    } // class Tests

    [Flags]
    public enum ShortPropTypeFlags: uint
    {
        MV = 0x00000001, // PDTF_MULTIPLEVALUES
        IN = 0x00000002, // PDTF_ISINNATE
        GP = 0x00000004, // PDTF_ISGROUP
        GB = 0x00000008, // PDTF_CANGROUPBY
        SB = 0x00000010, // PDTF_CANSTACKBY
        TP = 0x00000020, // PDTF_ISTREEPROPERTY
        FT = 0x00000040, // PDTF_INCLUDEINFULLTEXTQUERY
        VW = 0x00000080, // PDTF_ISVIEWABLE
        QR = 0x00000100, // PDTF_ISQUERYABLE
        CP = 0x00000200, // PDTF_CANBEPURGED
        SW = 0x00000400, // PDTF_SEARCHRAWVALUE
        SY = 0x80000000, // PDTF_ISSYSTEMPROPERTY
    }

}


// Look up non-existent property key in propstore should result in null.