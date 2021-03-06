﻿namespace XEws.CmdletAbstract
{
    using System;
    using System.Collections.Generic;
    using System.Management.Automation;
    using Microsoft.Exchange.WebServices.Data;

    /// <summary>
    /// Represent abstract class for cmdlets Get-XEwsFolder*
    /// </summary>
    public abstract class XEwsFolderCmdlet : XEwsCmdlet
    {
        #region Cmdlet parameters

        [Parameter(Mandatory = false, Position = 0)]
        public Folder FolderRoot
        {
            get;
            set;
        }

        [Parameter(Mandatory = false, Position = 2)]
        public SwitchParameter Recurse
        {
            get;
            set;
        }

        #endregion


        #region Methods

        /// <summary>
        /// Traverse through all folders within specified search root.
        /// </summary>
        /// <param name="searchRoot">Root of the search.</param>
        /// <param name="foundFolders">Reference variable where result will be returned.</param>
        internal void GetRecurseFolder(Folder searchRoot, ref List<Folder> foundFolders)
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            if (searchRoot.ChildFolderCount == 0)
                return;

            FindFoldersResults foundFolderResult = ewsSession.FindFolders(searchRoot.Id, new FolderView(100));

            foreach (Folder folder in foundFolderResult)
            {
                GetRecurseFolder(folder, ref foundFolders);
                foundFolders.Add(folder);
            }
        }

        /// <summary>
        /// Method is returning Folder object based on the search name.
        /// </summary>
        /// <param name="searchRoot">Folder from where search should begin.</param>
        /// <param name="folderName">Name of the folder to search.</param>
        /// <returns></returns>
        internal Folder GetFolder(Folder searchRoot, string folderName)
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            List<Folder> folders = GetFolder(searchRoot);

            foreach (Folder folder in folders)
            {
                if (folder.DisplayName == folderName)
                    return folder;
            }

            return null;
        }

        /// <summary>
        /// Method is returning List of the Folder object under search root.
        /// </summary>
        /// <param name="searchRoot">Folder from where search should begin.</param>
        /// <returns></returns>
        internal List<Folder> GetFolder(Folder searchRoot)
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            List<Folder> folders = new List<Folder>();
            FindFoldersResults findResult = ewsSession.FindFolders(searchRoot.Id, new FolderView(100));

            foreach (Folder folder in findResult)
                folders.Add(folder);

            return folders;
        }

        /// <summary>
        /// Method is returning Folder object from WellKnownFolderName enum.
        /// </summary>
        /// <param name="wellKnownFolderName">WellKnownFolderName for which Folder object needs to be returned.</param>
        /// <param name="ewsSession">ExchangeService session context.</param>
        /// <returns></returns>
        public static Folder GetWellKnownFolder(WellKnownFolderName wellKnownFolderName, ExchangeService ewsSession)
        {
            FolderId wellKnownFolderId = new FolderId(wellKnownFolderName);
            Folder wellKnownFolder = Folder.Bind(ewsSession, wellKnownFolderId);

            return wellKnownFolder;
        }

        /// <summary>
        /// Method is creating new folder under specified root folder.
        /// </summary>
        /// <param name="folderName">Name of the folder to create.</param>
        /// <param name="folderRoot">Root folder under which new folder will be created.</param>
        internal void AddFolder(string folderName, Folder folderRoot)
        {
            ExchangeService ewsSession = this.GetSessionVariable();

            Folder folder = new Folder(ewsSession);
            folder.DisplayName = folderName;
            folder.FolderClass = "IPF.Note";
            folder.Save(folderRoot.Id);
        }

        /// <summary>
        /// Method used to delete folder. If folder contains item or subfolder, force switch should be used.
        /// </summary>
        /// <param name="folder">Folder to delete.</param>
        /// <param name="deleteMode">Delete mode operation.</param>
        /// <param name="force">Overrides child folder and item check. </param>
        internal void RemoveFolder(Folder folder, DeleteMode deleteMode, bool force)
        {
            if (force)
            {
                this.RemoveFolder(folder, deleteMode);
            }
            else
            {
                int childFolderCount = folder.ChildFolderCount;
                int folderItemCount = folder.TotalCount;

                if (childFolderCount == 0 && folderItemCount == 0)
                    this.RemoveFolder(folder, deleteMode);
                else
                {
                    throw new InvalidOperationException(String.Format("Folder not empty or contains subfolders. ChildFolders: {0}, ItemsInFolder: {1}", childFolderCount, folderItemCount));
                }
            }
        }

        /// <summary>
        /// Helper method for deleting folder.
        /// </summary>
        /// <param name="folder">Folder to delete.</param>
        /// <param name="deleteMode">Delete mode operation.</param>
        private void RemoveFolder(Folder folder, DeleteMode deleteMode)
        {
            folder.Delete(deleteMode);
        }

        /// <summary>
        /// Method is moving folder to another location.
        /// </summary>
        /// <param name="folderToMove">Folder which will be moved.</param>
        /// <param name="destinationFolder">New destination.</param>
        internal void MoveFolder(Folder folderToMove, Folder destinationFolder)
        {
            // No need to check if folder to move is WellKnownFolder since those cannot be moved anyway.

            Folder checkFolderToMove = this.GetFolder(destinationFolder, folderToMove.DisplayName);

            if (checkFolderToMove != null)
                throw new InvalidOperationException(String.Format("Folder '{0}' already exist under '{1}'.", folderToMove.DisplayName, destinationFolder.DisplayName));

            folderToMove.Move(destinationFolder.Id);
        }

        #endregion
    }
}
