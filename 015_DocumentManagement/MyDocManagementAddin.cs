using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Atys.PowerEDIT;
using Atys.PowerEDIT.Extensibility;

namespace TeamSystem.Customizations
{
    [AddinData("MyDocManagementAddin", "PowerEDIT Addin",
        "24CFAE81-AA60-43E0-8AF4-BF14B934A15A",
        "1.0.0", "TeamSystem", "Dev", true, false, false, "", ImageFilename = "TS_LogoSmall_32.png")]
    public class MyDocManagementAddin : IPowerEDITAddin
    {
        #region const

        //caption barre
        private const string DOCUMENTMANAGEMENTCAPTION = "Document management Bar";

        //caption gruppi
        private const string DOCUMENTMANAGEMENTGROUPCAPTION = "Document management group";

        //caption pulsanti
        private const string REPLACEBUTTONCAPTION = "Replace";
        private const string NEWDOCUMENTBUTTONCAPTION = "New Doc";
        private const string ADDDATABUTTONCAPTION = "Add Data";
        private const string MOVEBUTTONCAPTION = "Move";
        private const string ONLYADMINBUTTONCAPTION = "Only Admin";

        #endregion

        #region fields

        private ExtensionState _AddinState = ExtensionState.Unknown;
        private IPowerEDITApp _PowerEDITApp = null;
        private IExtensionUIManager m_UIManager = null;
        private bool _IsUIManagerConnected = false;

        private UIMenuBarInfo _DocumentManagementBar = null;

        private UIMenuBarGroupInfo _DocumentManagementGroup = null;
        

        private UIMenuBarItemInfo _ReplaceButton = null;
        private UIMenuBarItemInfo _NewDocumentButton = null;
        private UIMenuBarItemInfo _MoveButton = null;
        private UIMenuBarItemInfo _AddDataButton = null;
        private UIMenuBarItemInfo _OnlyAdminButton = null;

        #endregion

        /// <summary>
        /// Inizializza una nuova istanza della classe
        /// </summary>
        /// <remarks>Il costruttore deve essere senza parametri</remarks>
        public MyDocManagementAddin()
        {
#if DEBUG
            Debugger.Launch();
#endif
        }

        #region IPowerEDITExtension Members

        public void Initialize(IPowerEDITApp pweApp)
        {
            if (pweApp == null)
                throw new ArgumentNullException(nameof(pweApp));

            //eventuale controllo su stato non running

            this._AddinState = ExtensionState.Initailizing;
            this._PowerEDITApp = pweApp;
            
            this._DocumentManagementBar = new UIMenuBarInfo(this, DOCUMENTMANAGEMENTCAPTION, AddinMenuBarTarget.AddinCustomBar);
            this.m_UIManager.CreateMenuBar(this._DocumentManagementBar);

            this._DocumentManagementGroup = new UIMenuBarGroupInfo(this, this._DocumentManagementBar, DOCUMENTMANAGEMENTGROUPCAPTION);
            this.m_UIManager.AppendGroupToMenuBar(this._DocumentManagementBar, this._DocumentManagementGroup);

            // New document button
            this._NewDocumentButton = new UIMenuBarItemInfo(this, this._DocumentManagementGroup,
                NEWDOCUMENTBUTTONCAPTION, AddinMenuItemType.Button,
                Properties.Resources.TS_LogoSmall_32);
            this.m_UIManager.AppendItemToMenuGroup(this._DocumentManagementGroup, this._NewDocumentButton);

            // Move button
            this._MoveButton = new UIMenuBarItemInfo(this, this._DocumentManagementGroup,
                MOVEBUTTONCAPTION, AddinMenuItemType.Button,
                Properties.Resources.TS_LogoSmall_32);
            this.m_UIManager.AppendItemToMenuGroup(this._DocumentManagementGroup, this._MoveButton);

            // Replace button
            this._ReplaceButton = new UIMenuBarItemInfo(this, this._DocumentManagementGroup,
                REPLACEBUTTONCAPTION, AddinMenuItemType.Button,
                Properties.Resources.ReplaceIcon);
            this.m_UIManager.AppendItemToMenuGroup(this._DocumentManagementGroup, this._ReplaceButton);
            
            // Add button
            this._AddDataButton = new UIMenuBarItemInfo(this, this._DocumentManagementGroup,
                ADDDATABUTTONCAPTION, AddinMenuItemType.Button,
                Properties.Resources.TS_LogoSmall_32);
            this.m_UIManager.AppendItemToMenuGroup(this._DocumentManagementGroup, this._AddDataButton);


            // Add button only for admin
            var powerDoc = this._PowerEDITApp.PowerDOCConnector;
            if (powerDoc.GetCurrentUser().AccessLevel == PowerDOCUserAccessLevel.Administrator)
            {
                this._OnlyAdminButton = new UIMenuBarItemInfo((IPowerEDITExtension) this, this._DocumentManagementGroup,
                    ONLYADMINBUTTONCAPTION, AddinMenuItemType.Button, Properties.Resources.TS_LogoSmall_32);
                this.m_UIManager.AppendItemToMenuGroup(this._DocumentManagementGroup, this._OnlyAdminButton);
            }

            this._AddinState = ExtensionState.Initialized;
        }

        public void Run()
        {
            Debug.Assert(this._PowerEDITApp != null);

            this._AddinState = ExtensionState.Running;

            //do something...

            this._PowerEDITApp.ActiveDocumentChanged += this._PowerEDITApp_ActiveDocumentChanged;

            this._PowerEDITApp.DocumentLoaded += this._PowerEDITApp_DocumentLoaded;
            this._PowerEDITApp.DocumentCreated += this._PowerEDITApp_DocumentCreated;
            this._PowerEDITApp.ClosingDocument += this._PowerEDITApp_ClosingDocument;
        }

        private void _PowerEDITApp_DocumentCreated(object sender, DocumentDataEventArgs e)
        {
            if (!e.IsPWEDoc)
                return;

            var pweDoc = (IPWEDoc) e.Document;

            pweDoc.TextChanging += this.PweDoc_TextChanging;
        }

        private void _PowerEDITApp_DocumentLoaded(object sender, DocumentDataEventArgs e)
        {
            if (!e.IsPWEDoc)
                return;

            var pweDoc = (IPWEDoc)e.Document;

            pweDoc.TextChanging += this.PweDoc_TextChanging;
        }

        private void _PowerEDITApp_ClosingDocument(object sender, DocumentDataEventArgs e)
        {
            if (!e.IsPWEDoc)
                return;

            var pweDoc = (IPWEDoc)e.Document;
            try
            {
                pweDoc.TextChanging -= this.PweDoc_TextChanging;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void PweDoc_TextChanging(object sender, TextChangeCancelEventArgs e)
        {
            if (e.LinesDeleted > 0 &&  e.DeletedText.Contains("non cancellare"))
            {
                e.Cancel = true;
            }
        }

        private void _PowerEDITApp_ActiveDocumentChanged(object sender, ActiveDocumentChangedEventArgs e)
        {
            if (e.PreviousDoc is IPWEDoc previousPweDoc)
                previousPweDoc.TextChanging -= this.PweDoc_TextChanging;
            
            if (e.ActiveDoc is IPWEDoc activePweDoc)
                activePweDoc.TextChanging += this.PweDoc_TextChanging;
            //aggancio e sgancio da eventi documento attivo in continuo
            //con stesso codice sopra
        }

        public void Shutdown()
        {
            Debug.Assert(this.m_UIManager != null);

            this._AddinState = ExtensionState.Closing;

            //UI
            if (this._DocumentManagementGroup != null)
                this.m_UIManager.DetachGroupFromMenuBar(this._DocumentManagementGroup);

            //this.m_UIManager.DetachGroupFromMenuBar(this._MesGroup);

            this.m_UIManager.DestroyMenuBar(this._DocumentManagementBar);
            //shut down code here...

            this._AddinState = ExtensionState.Closed;
        }

        #endregion

        #region IPowerEDITAddin Members

        /// <summary>
        /// Gestione notifica di pulsante dell'addin su ribbon premuto
        /// </summary>
        /// <param name="menuItem"></param>
        public void MenuButtonActionNotification(UIMenuBarItemInfo menuItem)
        {
            Debug.Assert(menuItem != null);

            var activeDoc = this._PowerEDITApp.GetPWEActiveDoc();
            
            if (menuItem.Caption == NEWDOCUMENTBUTTONCAPTION)
                NewDocumentAction();
            else if (menuItem.Caption == MOVEBUTTONCAPTION)
                MoveAction(activeDoc);
            else if (menuItem.Caption == REPLACEBUTTONCAPTION)
                ReplaceAction(activeDoc);
            else if (menuItem.Caption == ADDDATABUTTONCAPTION)
                AddDataAction(activeDoc);
        }

        private void NewDocumentAction()
        {
            // Creazione nuovo documento e salvataggio
            var newDoc = this._PowerEDITApp.CreatePWEDoc();
            //var x = this._PowerEDITApp.OpenPWEDoc(@"");
            newDoc.AddLine("Prova su nuovo documento");
            newDoc.SaveDocAs(@"");
            newDoc.Activate();
        }

        private void MoveAction(IPWEDoc activeDoc)
        {
            // Muoversi tra le righe
            activeDoc.MoveCursor(MoveCursorTarget.ToDocumentBegin);
            activeDoc.MoveCursor(MoveCursorTarget.PageDown);
            activeDoc.MoveCursor(MoveCursorTarget.PageUp);

            activeDoc.SetCursorPosition(1, 1);
            for (var i = 1; i <= activeDoc.LineCount; i++)
            {
                activeDoc.SetCursorPosition(i, 1);
                var line = activeDoc.GetLine(i);
                //...
                activeDoc.SetSelection(line.TextRange);
            }
        }

        private void ReplaceAction(IPWEDoc activeDoc)
        {
            // Find & Replace
            if (activeDoc == null)
            {
                MessageBox.Show("Documento attivo non è di tipo testo");
                return;
            }

            var profileName = this._PowerEDITApp.GetActiveProfileName();
            if (profileName != "PWEBase")
                return;

            const string stringToReplace = "S";
            const string replaceWith = "Raggio";

            // Ricerca e sostituzione testo MODO 1 (Manuale)
            for (var i = 1; i <= activeDoc.LineCount; i++)
            {
                var line = activeDoc.GetLine(i);
                var newLine = line.Text.Replace(stringToReplace, replaceWith);
                line.ReplaceAllText(newLine);
            }

            // Ricerca e sostituzione testo MODO 2 (API) - solo find
            /*
            var findResults = activeDoc.FindText(stringToReplace, SearchAction.SearchAll, SearchScope.All,
                SearchType.Standard, SearchCasing.Any,
                SearchWord.AnyText, string.Empty, true,
                UiInteraction.None);
            */

            // Ricerca e sostituzione testo MODO 3 (API) - replace diretto
            /*
            var replaceResults = activeDoc.ReplaceText(stringToReplace, replaceWith,
                SearchAction.SearchAll, SearchScope.All,
                SearchType.Standard, SearchCasing.Any,
                SearchWord.AnyText, string.Empty,
                UiInteraction.None);
            */
        }
        
        private void AddDataAction(IPWEDoc activeDoc)
        {
            activeDoc.MoveCursor(MoveCursorTarget.ToDocumentEnd);
            activeDoc.AddLine("nuova riga");
            //activeDoc.AddLineBefore();
        }

        #region standard implementation

        public ExtensionState AddinState
        {
            get { return this._AddinState; }
        }

        public bool IsUIManagerConnected
        {
            get { return this._IsUIManagerConnected; }
        }

        public void ConnectToUIManager(IExtensionUIManager uiManager)
        {
            if (uiManager == null)
                throw new ArgumentNullException();

            this.m_UIManager = uiManager;
            this._IsUIManagerConnected = true;

            //NB: inserire qui la costruzione dei componenti della UI?
        }

        public void DisconnectFromUIManager()
        {
            //NB: inserire qui la distruzione dei componenti della UI

            this.m_UIManager = null;
            this._IsUIManagerConnected = false;
        }

        public void MenuCheckedChangedNotification(UIMenuBarItemInfo menuItem, bool newState)
        {
            throw new NotImplementedException();
        }

        public void ReloadOptions()
        {
            throw new NotImplementedException();
        }

        public string OptionsPersistenceFilename
        {
            get { throw new NotImplementedException(); }
        }

        public object GetOptions()
        {
            throw new NotImplementedException();
        }

        public bool SetOptions(object options)
        {
            throw new NotImplementedException();
        }

        #endregion

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            //clean up code here...
        }

        #endregion
    }
}