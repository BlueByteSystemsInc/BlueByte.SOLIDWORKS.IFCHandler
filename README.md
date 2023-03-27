# BlueByte.SOLIDWORKS.IFCHandler
Requires xBIM.Ifc and only works with IFC4

# Sample


```

 static void Main(string[] args)
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

            handler.OpenDocument(editor, @"C:\Users\jlili\Desktop\Root.ifc");

            handler.BuildTree();

            handler.TraverseAndDo(x => Console.WriteLine($"{x.GetIndentation()} {x.Name}"));

            handler.TraverseAndEdit(x =>
            {
                x.UnsafeObject.Name = $"{x.Name} - XX";
                Console.WriteLine($"{x.GetIndentation()} {x.Name}");
            });

            handler.Dispose();

        }
        
        ```
