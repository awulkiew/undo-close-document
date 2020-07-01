//------------------------------------------------------------------------------
// <copyright>
//     Copyright (c) Adam Wulkiewicz.
// </copyright>
//------------------------------------------------------------------------------

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;

namespace UndoCloseDocument
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class UndoCloseDocument
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7fe92939-5f33-4bab-a8b7-f1322d158c5e");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private IVsUIShell uiShell;
        private DTE2 dte;

        private Events events;
        private SolutionEvents solutionEvents;
        //private DocumentEvents documentEvents;
        private WindowEvents windowEvents;
        //private CommandEvents commandEvents;

        private UndoCloseDocumentOptionPage optionPage;

        private List<Path> closedDocuments = new List<Path>();
        private HashSet<Path> openedDocuments = new HashSet<Path>(new PathComparer());

        class Path
        {
            public Path(string path)
            {
                Original = path;
                CaseInsensitive = path.ToLowerInvariant();
            }

            public bool IsSame(Path other)
            {
                return CaseInsensitive == other.CaseInsensitive;
            }

            public readonly string Original;
            public readonly string CaseInsensitive;
        }

        class PathComparer : IEqualityComparer<Path>
        {
            public int GetHashCode(Path p) { return p.CaseInsensitive.GetHashCode(); }
            public bool Equals(Path p, Path q) { return p.CaseInsensitive == q.CaseInsensitive; }
        }

        // See:
        // https://docs.microsoft.com/en-us/visualstudio/extensibility/dynamically-adding-menu-items
        class DynamicItemMenuCommand : OleMenuCommand
        {
            public static readonly int CommandId = 0x104;

            public DynamicItemMenuCommand(CommandID rootId, Predicate<int> matches, EventHandler invokeHandler, EventHandler beforeQueryStatusHandler)
                : base(invokeHandler, null /*changeHandler*/, beforeQueryStatusHandler, rootId)
            {
                this.matches = matches ?? throw new ArgumentNullException("matches");
            }

            public override bool DynamicItemMatch(int cmdId)
            {
                // Call the supplied predicate to test whether the given cmdId is a match.
                // If it is, store the command id in MatchedCommandid
                // for use by any BeforeQueryStatus handlers, and then return that it is a match.
                // Otherwise clear any previously stored matched cmdId and return that it is not a match.
                if (this.matches(cmdId))
                {
                    this.MatchedCommandId = cmdId;
                    return true;
                }

                this.MatchedCommandId = 0;
                return false;
            }

            private Predicate<int> matches;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UndoCloseDocument"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private UndoCloseDocument(UndoCloseDocumentPackage package, OleMenuCommandService commandService, IVsUIShell uiShell, DTE2 dte)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            this.uiShell = uiShell ?? throw new ArgumentNullException("IVsUIShell");
            this.dte = dte ?? throw new ArgumentNullException("DTE");

            this.events = this.dte.Events ?? throw new NullReferenceException(nameof(this.events));
            this.solutionEvents = this.events.SolutionEvents ?? throw new NullReferenceException("SolutionEvents");
            this.solutionEvents.Opened += SolutionEvents_Opened;
            //this.solutionEvents.BeforeClosing += SolutionEvents_BeforeClosing;
            this.solutionEvents.AfterClosing += SolutionEvents_AfterClosing;
            //this.documentEvents = this.events.DocumentEvents ?? throw new NullReferenceException("DocumentEvents");
            //this.documentEvents.DocumentOpening += DocumentEvents_DocumentOpening;
            //this.documentEvents.DocumentOpened += DocumentEvents_DocumentOpened;
            //this.documentEvents.DocumentClosing += DocumentEvents_DocumentClosing;
            this.windowEvents = this.events.WindowEvents ?? throw new NullReferenceException("WindowEvents");
            this.windowEvents.WindowCreated += WindowEvents_WindowCreated;
            this.windowEvents.WindowClosing += WindowEvents_WindowClosing;
            //this.commandEvents = this.events.CommandEvents ?? throw new NullReferenceException("CommandEvents");
            //this.commandEvents.BeforeExecute += CommandEvents_BeforeExecute;
            //this.commandEvents.AfterExecute += CommandEvents_AfterExecute;

            this.optionPage = package.GetDialogPage(typeof(UndoCloseDocumentOptionPage)) as UndoCloseDocumentOptionPage;
            if (this.optionPage == null)
                throw new NullReferenceException("OptionPage");

            var menuCommandID = new CommandID(CommandSet, CommandId);
            //var menuItem = new MenuCommand(this.Execute, menuCommandID);
            var menuItem = new OleMenuCommand(this.Execute, null, OnBeforeQueryStatus, menuCommandID);
            commandService.AddCommand(menuItem);

            CommandID dynamicItemRootId = new CommandID(CommandSet, DynamicItemMenuCommand.CommandId);
            DynamicItemMenuCommand dynamicMenuCommand = new DynamicItemMenuCommand(dynamicItemRootId,
                                                                                   IsValidDynamicItem,
                                                                                   OnInvokedDynamicItem,
                                                                                   OnBeforeQueryStatusDynamicItem);
            commandService.AddCommand(dynamicMenuCommand);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static UndoCloseDocument Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(UndoCloseDocumentPackage package)
        {
            OleMenuCommandService commandService = package.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            IVsUIShell uiShell = package.GetService(typeof(SVsUIShell)) as IVsUIShell;
            DTE2 dte = package.GetService(typeof(DTE)) as DTE2;
            
            Instance = new UndoCloseDocument(package, commandService, uiShell, dte);
        }

        /// <summary>
        /// This function is the callback called before the menu item is shown to the user.
        /// </summary>
        private void OnBeforeQueryStatus(object sender, EventArgs args)
        {
            OleMenuCommand matchedCommand = (OleMenuCommand)sender;
            matchedCommand.Enabled = closedDocuments.Count > 0;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            if (closedDocuments.Count <= 0)
                return;

            int lastIndex = closedDocuments.Count - 1;
            string fullName = closedDocuments[lastIndex].Original;
            closedDocuments.RemoveAt(lastIndex);

            //if (!dte.ItemOperations.IsFileOpen(fullName))
                dte.ItemOperations.OpenFile(fullName);
        }

        private void SolutionEvents_AfterClosing()
        {
            openedDocuments.Clear();
        }

        private void SolutionEvents_BeforeClosing()
        {
        }

        private void SolutionEvents_Opened()
        {
            HandleWindowCreate();
        }

        private void DocumentEvents_DocumentClosing(Document document)
        {
        }

        private void DocumentEvents_DocumentOpening(string documentPath, bool readOnly)
        {
        }

        private void DocumentEvents_DocumentOpened(Document document)
        {
        }

        private void WindowEvents_WindowClosing(Window window)
        {
            HandleWindowClose(window, true);
        }

        private void WindowEvents_WindowCreated(Window window)
        {
            HandleWindowCreate(window, true);
        }

        private void CommandEvents_AfterExecute(string guid, int id, object customIn, object customOut)
        {
            if (id == 658 // CloseDocument
             || id == 627 // CloseAllDocuments
             || id == 219 // CloseSolution
             || id == 223 // FileClose
             || id == 222 // FileOpen
             || id == 221 // FileNew
             || id == 451 // FileOpenFromWeb
             || id == 261 // Open
             || id == 217 // OpenProject
             || id == 450 // OpenProjectFromWeb
             || id == 315 // OpenProjectItem
             || id == 218 // OpenSolution
             || id == 199 // OpenWith
             || id == 343 // LoadUnloadedProject
             || id == 412 // ReloadProject
             || id == 344 // UnloadLoadedProject
             || id == 413 // UnloadProject
             )
            {
                if (new Guid(guid) == new Guid("{5EFC7975-14BC-11CF-9B2B-00AA00573819}")) // VSStd97CmdID
                {
                }
            }
            else if (id == 22) // CloseAllButPinned
            {
                if (new Guid(guid) == new Guid("{D63DB1F0-404E-4B21-9648-CA8D99245EC3}")) // VSStd11CmdID
                {
                }
            }
            else if (id == 528) // CloseAllButToolWindows
            {
                if (new Guid(guid) == new Guid("{712C6C80-883B-4AAD-B430-BBCA5256FA9D}")) // VSStd15CmdID
                {
                }
            }
            else if (id == 1650 // CloseAllButThis
                  || id == 1982 // CloseProject 
                  )
            {
                if (new Guid(guid) == new Guid("{1496A755-94DE-11D0-8C3F-00C04FC2AAE2}")) // VSStd2KCmdID
                {
                }
            }
        }

        private void CommandEvents_BeforeExecute(string guid, int id, object customIn, object customOut, ref bool cancelDefault)
        {
        }

        private bool IsValidDynamicItem(int commandId)
        {
            int index = commandId - DynamicItemMenuCommand.CommandId;
            return index >= 0 && index < Math.Min(this.optionPage.Show, closedDocuments.Count);
        }

        private void OnBeforeQueryStatusDynamicItem(object sender, EventArgs args)
        {
            DynamicItemMenuCommand matchedCommand = (DynamicItemMenuCommand)sender;

            if (closedDocuments.Count > 0)
            {
                matchedCommand.Enabled = true;
                matchedCommand.Visible = true;

                int index = matchedCommand.MatchedCommandId == 0
                          ? 0
                          : matchedCommand.MatchedCommandId - DynamicItemMenuCommand.CommandId;

                if (index >= 0 && index < closedDocuments.Count)
                {
                    int i = closedDocuments.Count - 1 - index;
                    matchedCommand.Text = (index + 1).ToString() + " "
                                        + ShortifyPath(closedDocuments[i].Original, 60);
                }
                else
                {
                    // just in case
                    matchedCommand.Enabled = false;
                    matchedCommand.Visible = false;
                    matchedCommand.Text = "Error";
                }
            }
            else
            {
                matchedCommand.Enabled = false;
                matchedCommand.Visible = true;
                matchedCommand.Text = "Empty";
            }

            // Clear the ID because we are done with this item.
            matchedCommand.MatchedCommandId = 0;
        }

        private void OnInvokedDynamicItem(object sender, EventArgs args)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DynamicItemMenuCommand invokedCommand = (DynamicItemMenuCommand)sender;

            if (!invokedCommand.Enabled)
                return;

            int index = invokedCommand.MatchedCommandId == 0
                          ? 0
                          : invokedCommand.MatchedCommandId - DynamicItemMenuCommand.CommandId;

            if (index >= 0 && index < closedDocuments.Count)
            {
                int i = closedDocuments.Count - 1 - index;

                string fullName = closedDocuments[i].Original;
                closedDocuments.RemoveAt(i);

                //if (!dte.ItemOperations.IsFileOpen(fullName))
                    dte.ItemOperations.OpenFile(fullName);
            }
        }

        string ShortifyPath(string path, int maxLength)
        {
            if (path.Length <= maxLength)
                return path;

            int firstIndex = path.IndexOfAny(new char[] {'\\', '/'});
            int lastIndex = path.LastIndexOfAny(new char[] { '\\', '/' });

            if (firstIndex < 0 || lastIndex < 0 || firstIndex >= lastIndex)
                return path;

            int i = firstIndex;
            while (i < lastIndex && firstIndex + path.Length - i > maxLength)
            {
                i = path.IndexOfAny(new char[] { '\\', '/' }, i + 1);
                if (i < 0)
                    return path; // unexpected
            }

            return path.Substring(0, firstIndex + 1) + "..." + path.Substring(i);
        }

        private void HandleWindowClose(Window window = null, bool useWindow = true)
        {
            if (useWindow && window != null && window.Document != null)
            {
                Path path = new Path(window.Document.FullName);
                openedDocuments.Remove(path); // result could by used here too

                if (closedDocuments.FindIndex(p => p.IsSame(path)) < 0)
                    closedDocuments.Add(path);
            }
            else
            {
                ForEachWindowFrame(delegate (string dm, DocumentState ds)
                {
                    if (ds != DocumentState.DocumentOpened)
                    {
                        Path path = new Path(dm);
                        if (openedDocuments.Remove(path))
                        {
                            if (closedDocuments.FindIndex(p => p.IsSame(path)) < 0)
                                closedDocuments.Add(path);
                        }
                    }
                });
            }

            // Store max N recently closed documents
            if (closedDocuments.Count > this.optionPage.Remember)
                closedDocuments.RemoveRange(0, closedDocuments.Count - this.optionPage.Remember);
        }

        private void HandleWindowCreate(Window window = null, bool useWindow = true)
        {
            if (useWindow && window != null && window.Document != null)
            {
                Path path = new Path(window.Document.FullName);
                openedDocuments.Add(path); // result could by used here too

                int i = closedDocuments.FindIndex(p => p.IsSame(path));
                if (i >= 0)
                    closedDocuments.RemoveAt(i);
            }
            else
            {
                ForEachWindowFrame(delegate (string dm, DocumentState ds)
                {
                    Path path = new Path(dm);
                    if (openedDocuments.Add(path))
                    {
                        int i = closedDocuments.FindIndex(p => p.IsSame(path));
                        if (i >= 0)
                            closedDocuments.RemoveAt(i);
                    }
                });        
            }
        }

        enum DocumentState { Unknown, DocumentClosed, DocumentOpened }

        class WindowFrame
        {
            public WindowFrame(string documentMoniker, DocumentState documentState)
            {
                DocumentMoniker = documentMoniker;
                DocumentState = documentState;
            }

            public readonly string DocumentMoniker;
            public readonly DocumentState DocumentState;
        }

        delegate void WindowFrameDelegate(string documentMoniker, DocumentState documentState);

        void ForEachWindowFrame(WindowFrameDelegate windowFrameDelegate)
        {
            IEnumWindowFrames windowsEnum;
            if (uiShell.GetDocumentWindowEnum(out windowsEnum) != VSConstants.S_OK)
                return;

            for (;;)
            {
                IVsWindowFrame[] frames = new IVsWindowFrame[10];
                uint count = 0;
                if (windowsEnum.Next(10, frames, out count) != VSConstants.S_OK)
                    break;

                for (uint i = 0; i < count; ++i)
                {
                    IVsWindowFrame w = frames[i];
                    object o = null;
                    if (w.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out o) == VSConstants.S_OK)
                    {
                        string dm = o as string;
                        if (dm != null)
                        {
                            DocumentState ds = DocumentState.Unknown;
                            o = null;
                            if (w.GetProperty((int)__VSFPROPID.VSFPROPID_DocData, out o) == VSConstants.S_OK)
                            {
                                if (o != null)
                                    ds = DocumentState.DocumentOpened;
                                else
                                    ds = DocumentState.DocumentClosed;
                            }

                            windowFrameDelegate(dm, ds);
                        }
                    }
                }

                if (count < 10)
                    break;
            }
        }
    }
}
