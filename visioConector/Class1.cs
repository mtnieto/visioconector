using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
namespace visioConector
{
    public class Class1
    {
        private const string Fipath = @"C:\Users\bruno.ibanez\Desktop\Reutilizacion\repos\maquinaEstados.vsdx";

        [Test]
        public void basicTest()
        {
            List<string> result = visioprueba.Processor.GetVisioShapesFromFile(Fipath);
            foreach (string res in result)
            {
                Console.WriteLine(res);
            }
        }
    }
}
