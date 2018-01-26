using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void SetupStub()
        {
            try
            {
                string startFolder = @"C:\Temp";
                List<string> myFolders = new List<string>()
            {
                "Kasper","Jesper","Jonathan"
            };
                foreach (var n in myFolders)
                {
                    var p = $"{startFolder}\\{n}";
                    Directory.CreateDirectory(p);
                    File.Create($"{p}\\{n}.avi");
                }
            }
            catch (Exception)
            {
                
            }
        }
    }
}
