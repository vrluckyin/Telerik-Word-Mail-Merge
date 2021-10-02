using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Telerik.Documents.Common.Model;
using Telerik.Documents.Media;

using Telerik.Documents.Primitives;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Flow.Model.Shapes;

namespace Telerik.Documents.Flow.MailMergeUtil.Model
{
    /// <summary>
    /// A class that holds parsed token and actual Run object which are returned by Telerik 
    /// </summary>
    public class PlaceholderToken
    {
        //holds value of actual token
        //For example,  "[City], [State] [Zip]"
        //Following Run will be returned, that will be saved into "Parent" property
        //[City], [State] [Zip]
        //while following placeholders have been created
        //[
        //City
        //],<Space> 
        //[
        //State
        //]
        //<Space>
        //[
        //Zip
        //]
        public Run Placeholder { get; private set; }
        public Run Parent { get; private set; }
        public string Text { get { return Placeholder.Text; } }
        public PlaceholderToken(Run run, string text)
        {
            Parent = run;
            Placeholder = run.Clone(run.Document);
            Placeholder.Text = text;
        }

        public override string ToString()
        {
            return Text;
        }

        public void Cleanup()
        {
            if (Parent != null && Parent.Paragraph != null)
            {
                Parent.Paragraph.Inlines.Remove(Parent);
            }
        }
    }
}
