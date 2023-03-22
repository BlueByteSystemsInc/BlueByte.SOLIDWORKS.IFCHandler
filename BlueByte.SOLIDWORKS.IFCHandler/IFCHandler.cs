using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;


/// <summary>
/// IFC Hanlder
/// </summary>
namespace BlueByte.SOLIDWORKS.IFCHandler
{
    /// <summary>
    /// IFC Component
    /// </summary>
    public class IFCComponent
    {

        /// <summary>
        /// Froms the assembly.
        /// </summary>
        /// <param name="c">The c.</param>
        /// <returns></returns>
        public static IFCComponent FromAssembly(IIfcElementAssembly c)
        {
            var i = new IFCComponent();
            i.Name = c.Name;
            i.UnsafeObject = c as IIfcElement;
            return i;
        }
        /// <summary>
        /// Froms the definition.
        /// </summary>
        /// <param name="c">The c.</param>
        /// <returns></returns>
        public static IFCComponent FromDefinition(IIfcObjectDefinition c)
        {
            var i = new IFCComponent();
            i.Name = c.Name;
            i.UnsafeObject = c as IIfcElement;
            return i;
        }

        /// <summary>
        /// Gets the indentation.
        /// </summary>
        /// <returns></returns>
        public string GetIndentation()
        {
            var str = new StringBuilder();

            var p = Parent;

            while (p != null)
            {
                str.Append(" ");
                p = p.Parent;
            }



            return str.ToString();
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the unsafe object.
        /// </summary>
        /// <value>
        /// The unsafe object.
        /// </value>
        public IIfcElement UnsafeObject { get; set; }

        /// <summary>
        /// Creates new name.
        /// </summary>
        /// <value>
        /// The new name.
        /// </value>
        public string NewName { get; set; }

        /// <summary>
        /// Gets or sets the parent.
        /// </summary>
        /// <value>
        /// The parent.
        /// </value>
        public IFCComponent Parent { get; set; }

        /// <summary>
        /// Gets or sets the children.
        /// </summary>
        /// <value>
        /// The children.
        /// </value>
        public List<IFCComponent> Children { get; set; } = new List<IFCComponent>();
    }

    /// <summary>
    /// IFC Handler
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public class IFCHandler : IDisposable
    {

        /// <summary>
        /// Gets or sets the root component.
        /// </summary>
        /// <value>
        /// The root component.
        /// </value>
        public IFCComponent RootComponent { get; set; }

        /// <summary>
        /// Gets or sets the unsafe object.
        /// </summary>
        /// <value>
        /// The unsafe object.
        /// </value>
        public IfcStore UnsafeObject { get; set; }

        /// <summary>
        /// Opens the document.
        /// </summary>
        /// <param name="editor">The editor.</param>
        /// <param name="fileName">Name of the file.</param>
        public void OpenDocument(XbimEditorCredentials editor, string fileName)
        {
            UnsafeObject = IfcStore.Open(fileName, editor);

        }
        /// <summary>
        /// Builds the tree.
        /// </summary>
        public void BuildTree()
        {
            var root = UnsafeObject.Instances.OfType<IIfcElementAssembly>().FirstOrDefault();
            RootComponent = IFCComponent.FromAssembly(root);
            var children = root.IsDecomposedBy.ToArray().SelectMany(x => x.RelatedObjects).ToArray();
            Traverse(root, null);
        }
        /// <summary>
        /// Traverses the and edit.
        /// </summary>
        /// <param name="doAction">The do action.</param>
        public void TraverseAndEdit(Action<IFCComponent> doAction)
        {

            var txn = UnsafeObject.BeginTransaction("Change Names");

            Action<IFCComponent> traverse = default(Action<IFCComponent>);

            traverse = (IFCComponent component) => {

                if (doAction != null)
                    doAction.Invoke(component);

                var children = component.Children;

                foreach (var child in children)
                    traverse(child);
            };


            traverse(RootComponent);


            txn.Commit();

            txn.Dispose();

            UnsafeObject.SaveAs(UnsafeObject.FileName);
        }


        /// <summary>
        /// Traverses the and do.
        /// </summary>
        /// <param name="doAction">The do action.</param>
        public void TraverseAndDo(Action<IFCComponent> doAction)
        {
            Action<IFCComponent> traverse = default(Action<IFCComponent>);

            traverse = (IFCComponent component) => {

                if (doAction != null)
                    doAction.Invoke(component);

                var children = component.Children;

                foreach (var child in children)
                    traverse(child);
            };


            traverse(RootComponent);
        }


        /// <summary>
        /// Traverses the specified object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="parent">The parent.</param>
        void Traverse(IIfcObjectDefinition obj, IFCComponent parent)
        {
            var _children = obj.IsDecomposedBy.ToArray().SelectMany(x => x.RelatedObjects).ToArray();

            var _component = IFCComponent.FromDefinition(obj);
            _component.Parent = parent;
            if (parent != null)
                parent.Children.Add(_component);
            else
                RootComponent = _component;

            foreach (var _child in _children)
            {
                Traverse(_child, _component);
            }

        }


        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            UnsafeObject.Dispose();

        }



    }


}
