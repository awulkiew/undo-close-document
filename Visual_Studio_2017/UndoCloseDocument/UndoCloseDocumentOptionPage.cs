//------------------------------------------------------------------------------
// <copyright>
//     Copyright (c) Adam Wulkiewicz.
// </copyright>
//------------------------------------------------------------------------------

using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;

namespace UndoCloseDocument
{
    public class UndoCloseDocumentOptionPage : DialogPage
    {
        private int remember = 10;
        private int show = 10;

        [Category("General")]
        [DisplayName("Remember")]
        [Description("Maximum number of remembered documents.")]
        public int Remember
        {
            get { return remember; }
            set
            {
                if (value < 1)
                    remember = 1;
                else if (value > 100)
                    remember = 100;
                else
                    remember = value;

                if (remember < show)
                    show = remember;
            }
        }

        [Category("General")]
        [DisplayName("Show")]
        [Description("Maximum number of documents shown on the list in the menu Window.")]
        public int Show
        {
            get { return show; }
            set
            {
                if (value < 1)
                    show = 1;
                else if (value > 100)
                    show = 100;
                else
                    show = value;

                if (remember < show)
                    remember = show;
            }
        }
    }
}

