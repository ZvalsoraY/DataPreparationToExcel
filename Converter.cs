using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataPreparationToExcel
{
    static class Converter
    {
        //string fileName = $"20tfiEMAT_20_ver_st3_WT10_sens3p5";
        static public List<string> list = new List<string>();
        
        static public void consoleWriteCheck(List<string> list)
        {
            foreach (var s in list)
                Console.Write(" " + s);
            //Console.Write(s);
            Console.WriteLine();
        }
        
        static void consoleWriteCheck(List<List<double>> writingArray)
        {
            foreach (var item in writingArray)
            {
                foreach (var s in item)
                    Console.Write(" " + s);
                //Console.Write(s);
                Console.WriteLine();
            }
        }

    }
}
