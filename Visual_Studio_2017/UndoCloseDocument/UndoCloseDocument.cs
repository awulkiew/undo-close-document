﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

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

        private DTE2 dte;
        private Events events;
        private WindowEvents windowEvents;

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
        private UndoCloseDocument(Package package, OleMenuCommandService commandService, DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            this.dte = dte ?? throw new ArgumentNullException(nameof(dte));

            this.events = this.dte.Events ?? throw new NullReferenceException(nameof(this.events));

            // Use WindowEvents instead of DocumentEvents because
            //   document == null is passed into the the latter for
            //   existing but non previously viewed tags.
            this.windowEvents = this.events.WindowEvents ?? throw new NullReferenceException(nameof(this.windowEvents));
            this.windowEvents.WindowCreated += WindowEvents_WindowCreated;
            this.windowEvents.WindowClosing += WindowEvents_WindowClosing;

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
            ThreadHelper.ThrowIfNotOnUIThread();

            OleMenuCommandService commandService = package.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            DTE2 dte = package.GetService(typeof(DTE)) as DTE2;

            Instance = new UndoCloseDocument(package, commandService, dte);
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
            ThreadHelper.ThrowIfNotOnUIThread();

            if (closedDocuments.Count <= 0)
                return;

            int lastIndex = closedDocuments.Count - 1;
            string fullName = closedDocuments[lastIndex];
            closedDocuments.RemoveAt(lastIndex);

            foreach (Window w in dte.Windows)
                if (w.Document != null
                        && w.Document.FullName == fullName)
                    return; // already opened

            //if (!dte.ItemOperations.IsFileOpen(fullName))
                dte.ItemOperations.OpenFile(fullName);
        }

        private void WindowEvents_WindowCreated(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // All of this HashSet/loop magic is here because
            // neither dte.Events.DocumentEvents nor dte.Documents
            // nor window.Document are fully reliable.
            foreach (Window w in dte.Windows)
                if (w.Document != null
                        && openedDocuments.Add(w.Document.FullName))
                    closedDocuments.Remove(w.Document.FullName);

            if (window.Document != null)
                closedDocuments.Remove(window.Document.FullName);
        }

        private void WindowEvents_WindowClosing(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // All of this HashSet/loop magic is here because
            // neither dte.Events.DocumentEvents nor dte.Documents
            // nor window.Document are fully reliable.
            HashSet<string> docs = new HashSet<string>();
            foreach (Window w in dte.Windows)
                if (w.Document != null)
                    docs.Add(w.Document.FullName);

            openedDocuments.ExceptWith(docs);

            foreach (string s in openedDocuments)
                closedDocuments.Add(s);

            openedDocuments = docs;

            if (closedDocuments.Count > 10)
                closedDocuments.RemoveRange(0, closedDocuments.Count - 10);
        }

        private bool IsValidDynamicItem(int commandId)
        {
            int index = commandId - DynamicItemMenuCommand.CommandId;
            return index >= 0 && index < closedDocuments.Count;
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
                                        + ShortifyPath(closedDocuments[i], 60);
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

                string fullName = closedDocuments[i];
                closedDocuments.RemoveAt(i);

                foreach (Window w in dte.Windows)
                    if (w.Document != null
                            && w.Document.FullName == fullName)
                        return; // already opened

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

        int FindFirstOf(int i1, int i2)
        {
            int result = i1;
            if (i2 >= 0 && i2 < i1)
                result = i2;
            return result;
        }

        HashSet<string> openedDocuments = new HashSet<string>();
        List<string> closedDocuments = new List<string>();
    }
}