using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using SRL.Mappings.Enginering;
using SRL.Mappings.Enginering.PhysicalLevel;
using SRL.ResourceShape;
namespace visioConector
{
    public class Class1
    {
        private const string Fipath = @"C:\Users\User\Desktop\visioconector\maquinaEstados.vsdx";
        private const string Fipath2 = @"C:\Users\User\Desktop\visioconector\maquinaEstados2.vsdx";


        [Test]
        public void basicTest()
        {
            List<PhysicalModel> result = visioprueba.Processor.GetVisioShapesFromFile(Fipath);
            visioprueba.Processor.ParseToVisio(Fipath2, result);
        }
    }
}
