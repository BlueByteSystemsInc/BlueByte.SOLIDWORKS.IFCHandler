using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.Ifc;
using SolidWorks.Interop.sldworks;


namespace BlueByte.SOLIDWORKS.IFCHandler.Sandbox
{
    class Program
    {
        static void Main(string[] args)
        {
            SolidworksInterface swInterface = new SolidworksInterface();
            
            // Must have part pre-selected in the feature tree so this won't fail 
            var swComp = swInterface.GetSelectedComponent();

            // Select the commented part in the assembly and un-comment only its line to run/verify this.

            //var isSameComponent = swInterface.VerifyIFC_ComponentData(swComp, 0.12463, -0.32757, -0.100053); // Bigbag haak_copy<1>
            //var isSameComponent = swInterface.VerifyIFC_ComponentData(swComp, -.35444, .09776, -0.100053); // Bigbag haak_copy<5>
            //var isSameComponent = swInterface.VerifyIFC_ComponentData(swComp, .129774, -0.1, 0.09491); // Verloop DN150 naar blaaspot_copy<4>
            //var isSameComponent = swInterface.VerifyIFC_ComponentData(swComp, .125508, -.153956, -0.010974); // Blaaspot knevelstang<1>
            //var isSameComponent = swInterface.VerifyIFC_ComponentData(swComp, -.22, -.10955, -.10955); // Bladder Pipe_Copy<1> 
            var isSameComponent = swInterface.VerifyIFC_ComponentData(swComp, .16731, .055072, .185758); // 69942<1>


            Console.WriteLine("Component Match? " + isSameComponent);
            swInterface.Dispose();
            Console.ReadKey();
        }

        private void Example()
        {
            var handler = new IFCHandler();
            var editor = new XbimEditorCredentials
            {
                ApplicationDevelopersName = "Blue Byte Systems Inc.",
                ApplicationFullName = "IFC Exporter for SOLIDWORKS",
                ApplicationIdentifier = "IFC Exporter for SOLIDWORKS",
                ApplicationVersion = "4.0",
                EditorsFamilyName = "Jlili",
                EditorsGivenName = "Amen",
                EditorsOrganisationName = "Blue Byte Systems Inc."
            };

            handler.OpenDocument(editor, @"");

            handler.BuildTree();

            handler.TraverseAndDo(x => Console.WriteLine($"{x.GetIndentation()} {x.Name}"));

            handler.TraverseAndEdit(x =>
            {
                x.UnsafeObject.Name = $"{x.Name} - XX";
                Console.WriteLine($"{x.GetIndentation()} {x.Name}");
            });

            handler.Dispose();


            Console.Read();
        }
    }
}
