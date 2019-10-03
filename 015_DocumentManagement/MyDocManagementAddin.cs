using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
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
        private const string MYBARCAPTION = "My Addin Bar";

        //caption gruppi
        private const string MYGROUP1CAPTION = "Document management";

        //caption pulsanti
        private const string MYCOMMAND1BUTTONCAPTION = "Replace";

        #endregion

        #region fields

        private ExtensionState _AddinState = ExtensionState.Unknown;
        private IPowerEDITApp _PowerEDITApp = null;
        private IExtensionUIManager m_UIManager = null;
        private bool _IsUIManagerConnected = false;

        private UIMenuBarInfo _MyAdvancedBar = null;

        private UIMenuBarGroupInfo _MyGroup1 = null;
        private UIMenuBarGroupInfo _MyHomeGroup = null;

        private UIMenuBarItemInfo _MyCommandButton1 = null;
        private UIMenuBarItemInfo _MyCommandButton2 = null;

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

            //inizializzazioni varie...

            //UI

            //prendo il riferimento alla barra principale (Power)
            
            //aggiungo una barra per comandi speciali
            this._MyAdvancedBar = new UIMenuBarInfo(this, MYBARCAPTION, AddinMenuBarTarget.AddinCustomBar);
            this.m_UIManager.CreateMenuBar(this._MyAdvancedBar);

            //
            this._MyGroup1 = new UIMenuBarGroupInfo(this, this._MyAdvancedBar, MYGROUP1CAPTION);
            this.m_UIManager.AppendGroupToMenuBar(this._MyAdvancedBar, this._MyGroup1);

            this._MyCommandButton1 = new UIMenuBarItemInfo(this, this._MyGroup1,
                MYCOMMAND1BUTTONCAPTION, AddinMenuItemType.Button,
                Properties.Resources.ReplaceIcon);
            this.m_UIManager.AppendItemToMenuGroup(this._MyGroup1, this._MyCommandButton1);


            this._AddinState = ExtensionState.Initialized;
        }

        public void Run()
        {
            Debug.Assert(this._PowerEDITApp != null);

            this._AddinState = ExtensionState.Running;

            //do something...
        }

        public void Shutdown()
        {
            Debug.Assert(this.m_UIManager != null);

            this._AddinState = ExtensionState.Closing;

            //UI
            if (this._MyGroup1 != null)
                this.m_UIManager.DetachGroupFromMenuBar(this._MyGroup1);

            //this.m_UIManager.DetachGroupFromMenuBar(this._MesGroup);

            this.m_UIManager.DestroyMenuBar(this._MyAdvancedBar);
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

            if (menuItem.Caption == MYCOMMAND1BUTTONCAPTION)
            {
                var activeDoc = this._PowerEDITApp.GetTextActiveDoc();

                var profileName = this._PowerEDITApp.GetActiveProfileName();
                if (profileName != "PWEBase")
                    return;
                
                var countDoc = activeDoc.LineCount;
                for (var i = 1; i < countDoc; i++)
                {
                    var line = activeDoc.GetLine(i);

                    var stringToReplace = "S";
                    var replaceWith = "Raggio";

                    var newLine = line.Text.Replace(stringToReplace, replaceWith);
                    line.ReplaceAllText(newLine);
                }
            }

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