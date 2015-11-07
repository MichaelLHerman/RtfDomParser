using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace RtfDomParser.Tests
{
    [TestFixture]
    public class Tests
    {
        [Test]
        public void TestLoadRTFText()
        {
            var document = new RTFDomDocument();
            document.LoadRTFText(@"{\rtf1\fbidis\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Segoe UI;}{\f1\fnil Segoe UI;}}
{\colortbl ;\red0\green0\blue0;\red255\green0\blue0;\red0\green0\blue255;\red0\green128\blue0;}
{\*\generator Riched20 6.3.9600}\viewkind4\uc1 
\pard\ltrpar\tx720\cf1\f0\fs28 black\cf2 red\cf3 blue\cf4 green\par
green\b and\cf1 bold\b0 andnot\cf4\f1\par
}
");
            CollectionAssert.IsNotEmpty(document.Elements);
            

        }

        [Test]
        public void TestFiles()
        {

            foreach (var file in Assembly.GetExecutingAssembly().GetManifestResourceNames().Where(r => r.Contains("TestData") && r.EndsWith(".rtf")))
            {
                var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(file);
                var document = new RTFDomDocument();
                document.Load(stream);
            } 

        }
    }
}
