using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
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
    [AddinData("MyUseCasesAddin", "PowerEDIT Addin",
        "33449847-3710-4E36-93E6-5A1266081725",
        "1.0.0", "TeamSystem", "Dev", true, false, false, "", ImageFilename = "TS_LogoSmall_32.png")]
    public class MyUseCasesAddin : IPowerEDITAddin
    {
        #region fields

        private ExtensionState _AddinState = ExtensionState.Unknown;
        private IPowerEDITApp _PowerEDITApp = null;
        private IPowerDOCConnector _PowerDoc = null;
        private IExtensionUIManager m_UIManager = null;
        private bool _IsUIManagerConnected = false;

        #endregion

        /// <summary>
        /// Inizializza una nuova istanza della classe
        /// </summary>
        /// <remarks>Il costruttore deve essere senza parametri</remarks>
        public MyUseCasesAddin()
        {
            //do nothing
#if DEBUG
            System.Diagnostics.Debugger.Launch();
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
            this._PowerDoc = pweApp.PowerDOCConnector;
            
            this._AddinState = ExtensionState.Initialized;
        }

        public void Run()
        {
            Debug.Assert(this._PowerEDITApp != null);

            this._AddinState = ExtensionState.Running;

            this._PowerDoc.CreatingFindFormCustomFieldsStyles += OnCreatingFindFormCustomFieldsStyles;
            this._PowerDoc.BeforeShowingFindFormContextMenuItems += OnBeforeShowingFindFormContextMenuItems;
            this._PowerDoc.BeforeDncManualRxOperation += OnBeforeDncManualRxOperation;
            this._PowerDoc.BeforeDncTxOperation += OnBeforeDncTxOperation;
            this._PowerDoc.DncManualRxOperationCompleted += OnDncManualRxOperationCompleted;
            this._PowerDoc.CustomCommandsPanelRequested += OnCustomCommandsPanelRequested;
        }
        
        public void Shutdown()
        {
            Debug.Assert(this.m_UIManager != null);

            this._AddinState = ExtensionState.Closing;
            
            this._PowerDoc.CreatingFindFormCustomFieldsStyles -= OnCreatingFindFormCustomFieldsStyles;
            this._PowerDoc.BeforeShowingFindFormContextMenuItems -= OnBeforeShowingFindFormContextMenuItems;
            this._PowerDoc.BeforeDncManualRxOperation -= OnBeforeDncManualRxOperation;
            this._PowerDoc.BeforeDncTxOperation -= OnBeforeDncTxOperation;

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
            throw new NotImplementedException();
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

        #region Colorazione campi in base a condizioni

        private void OnCreatingFindFormCustomFieldsStyles(object sender, PowerDOCGridStylesEventArgs e)
        {
            const string fieldName = "StatoCorrente";
            e.GridStyles = new List<PowerDOCStyleCondition>
            {
                CreateStyleCondition(fieldName, Color.LightBlue, 1),
                CreateStyleCondition(fieldName, Color.LightCoral, 2),
                CreateStyleCondition(fieldName, Color.LightGreen, 3)
            };
        }

        #endregion

        #region Abilitazione menù

        private PowerDOCStyleCondition CreateStyleCondition(string fieldName, Color color, int value)
        {
            return new PowerDOCStyleCondition(fieldName, color, false, PowerDOCStyleConditionType.Expression,
                $"[{fieldName}] = {value}", null, null);
        }

        private void OnBeforeShowingFindFormContextMenuItems(object sender, PowerDOCContextMenuAuthorizationCancelEventArgs e)
        {
            if (e.CurrentUser.AccessLevel == PowerDOCUserAccessLevel.Administrator)
                return;

            var menuItems = e.MenuItems;

            foreach (var menuItem in menuItems)
            {
                switch (menuItem.MenuItemType)
                {
                    case PowerDOCContextMenuItemType.LockForNC:
                    case PowerDOCContextMenuItemType.UnlockForNC:
                        menuItem.Enabled = false;
                        break;
                    case PowerDOCContextMenuItemType.CheckIn:
                    case PowerDOCContextMenuItemType.CheckOut:
                    case PowerDOCContextMenuItemType.UndoCheckOut:
                    case PowerDOCContextMenuItemType.OpenSelected:
                    case PowerDOCContextMenuItemType.SendToDnc:
                    case PowerDOCContextMenuItemType.ManualRxFromDnc:
                    case PowerDOCContextMenuItemType.AddAttachment:
                    case PowerDOCContextMenuItemType.DeleteRow:
                    case PowerDOCContextMenuItemType.Images:
                    case PowerDOCContextMenuItemType.CopyFromRx:
                    case PowerDOCContextMenuItemType.DeleteFromRxPath:
                    case PowerDOCContextMenuItemType.EditRow:
                    case PowerDOCContextMenuItemType.DeleteAttachment:
                    case PowerDOCContextMenuItemType.AdminCommandsGroup:
                    case PowerDOCContextMenuItemType.AdminForceCheckIn:
                    case PowerDOCContextMenuItemType.AdminForceUndoCheckOut:
                    case PowerDOCContextMenuItemType.CustomerCommand1:
                    case PowerDOCContextMenuItemType.CustomerCommand2:
                    case PowerDOCContextMenuItemType.CustomerCommand3:
                    case PowerDOCContextMenuItemType.CustomerCommand4:
                    case PowerDOCContextMenuItemType.LoadAttachmentsPreview:
                    case PowerDOCContextMenuItemType.ReviewsGroup:
                    case PowerDOCContextMenuItemType.ReviewsRestore:
                    case PowerDOCContextMenuItemType.ReviewsStore:
                    case PowerDOCContextMenuItemType.ReviewsCompare:
                    case PowerDOCContextMenuItemType.ReviewsCompareToCurrent:
                    case PowerDOCContextMenuItemType.ReviewsDelete:
                    case PowerDOCContextMenuItemType.Print:
                    case PowerDOCContextMenuItemType.ShowHistory:
                    case PowerDOCContextMenuItemType.OpenAttachment:
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        private void OnCustomCommandsPanelRequested(object sender, PowerDOCFormReferenceEventArgs e)
        {
            var currentRow = this._PowerDoc.GetRecordById(e.IdData);

            //Verifico che la condizione sia soddisfatta
            var macchina = currentRow.Field<string>("Macut");
            var percorso = currentRow.Field<string>("Percorso");
            var codice = currentRow.Field<string>("Codice");
            

            const string target = @"C:\Cnc\Backup";
            var fileName = Path.GetFileName(percorso);
            if (fileName != null)
                File.Copy(percorso, Path.Combine(target, macchina, fileName));
        }

        #endregion

        #region Controllo trasmissione programmi

        /*
         * TX
         * Creazione part program
         * Lista utensili
         * Presetting (creazione allegato)
         *  - Stato 2 = non è stato fatto il presetting
         *
         * RX
         */

        private void OnBeforeDncTxOperation(object sender, PowerDOCTxOperationSetupCancelEventArgs e)
        {
            //Recupero il record corrente in base all’IdData del record
            var currentRow = this._PowerDoc.GetRecordById(e.IdData);

            //Verifico che la condizione sia soddisfatta
            if (currentRow.Field<int>("StatoCorrente") == 2)
            {
                MessageBox.Show("Programma bloccato",
                    "Controllo stato",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                //Annullo la trasmissione
                e.Cancel = true;
            }
        }

        private void OnBeforeDncManualRxOperation(object sender, PowerDOCRxOperationSetupCancelEventArgs e)
        {
            //Recupero il record corrente in base all’IdData del record
            var currentRow = this._PowerDoc.GetRecordById(e.IdData);

            //Verifico che la condizione sia soddisfatta
            if (currentRow.Field<int>("Stato") == 3)
            {
                MessageBox.Show("Programma bloccato",
                    "Controllo stato",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                //Annullo la trasmissione
                e.Cancel = true;
            }
        }

        private void OnDncManualRxOperationCompleted(object sender, PowerDOCRxOperationResultEventArgs e)
        {
            var fileFullPath = e.DocumentFullPath;
            var rxFolder = e.RxFolder;
        }

        #endregion
    }
}