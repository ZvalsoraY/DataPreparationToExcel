using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using DataPreparationToExcelNS;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using NUnit.Framework;
using Assert = NUnit.Framework.Assert;

namespace DataPreparationToExcelNS.test
{
    
    [TestClass]
    public class ConverterConverterToExcelTest
    {
        [TestMethod]
        public void FileDoubleArrayListTest()
        {
            IEnumerable<IEnumerable> expectedRes = new IEnumerable[] 
            { new double[] { 1.2, 2.3, 3.4, 4.5, 5.6, 6.7, 7.8 },
              new double[] { 2.3, 3.4, 4.5, 5.6, 6.7, 7.8, 8.9 },
              new double[] { 3.4, 4.5, 5.6, 6.7, 7.8, 8.9, 9.9 }};

            var result = ConverterToExcel.FileDoubleArrayList("inputFileTest.txt");
            Assert.That(expectedRes, Is.EqualTo(result).AsCollection.Within(1.0E-7));
            
        }
    }
}
