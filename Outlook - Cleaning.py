import os, pandas as pd, datetime, re, sys, traceback, time, logging
import win32com.client, pythoncom #pip install pywin32
from win32com.client import constants as winConstants
from PySide6 import QtCore, QtWidgets, QtGui
"""This program scans all your local outlook personal mail box, to find duplicate messages. See HELP_TEXT below."""

FOLDER_SIZE_LIMIT = 20000 #Can be adapted
EXCLUDED_FOLDERS = [] #Here you may exclude the scanning of some folders, by their name. Ex: ['Boîte de réception',]
INCLUDE_SUBFOLDERS = True #You might want to scan only top directory
LOGGING_LEVEL = logging.DEBUG #DEBUG, INFO, WARNING, ERROR, CRITICAL
VERSION = "2022-09-16"
HELP_TEXT = """This program will:

1) scan your local outlook mail boxes
- This can be you personal mail, or a shared mailbox if you have some
- It does not scan old messages that are in the cloud

2) Propose you some emails for deletion
- The following emails are proposed:
    - Receipts for undelivered or recalled mails. Unless they are still unread.
    - Mails whose content (text, and attachments names) is contained in a more recent message of the same conversation.
        For example: You replied to a message. It will propose you to keep your reply and delete the initial message,
        unless it contains attachments that were not in your reply.
    - It won't delete a mail if the newer message is in your Deleted folder. Unless the initial mail is also in the Deleted folder
- You will have the choice to accept or refuse each proposal
- You will see in the status bar the total size of the selected mails
- Selected emails will be moved to your Deleted Items folder. To delete them permanently, relaunch the scan.

We recomend to launch this program when you do not need Outlook:
Depending on the size of your box, it may run 10 minutes or more, during which outlook will be very slow."""

class WorkerSignals(QtCore.QObject): #Generic class defining signals available for a Worker in a thread
    finished = QtCore.Signal()
    error = QtCore.Signal(tuple)
    newStep = QtCore.Signal(str, int, int) #specific signal
    result = QtCore.Signal(object)
    progress = QtCore.Signal(int) #will be implemented via a callback function

class Worker(QtCore.QRunnable): #Generic class to build a worker. Will be used for all backend operations on mails
    def __init__(self, fn, *args, **kwargs): #Create a worker to launch fn in a thread
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals() #Define signals
        self.kwargs['progress_callback'] = self.signals.progress #Add the callback to track progress to our kwargs

    @QtCore.Slot()
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs) #Call fn, with kwargs (including the callback for progress)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)  #Return the result of the processing
        finally:
            self.signals.finished.emit()

class ui_MainWindow(QtWidgets.QMainWindow): #Bare user interface without any logic attached to widgets
    def __init__(self, ):
        super().__init__()
        self.setWindowTitle(f"Outlook Duplicates Cleaner - version {VERSION}")
        menu = self.menuBar()
        self.help_action = QtGui.QAction("&Help", self)
        menu.addAction(self.help_action)

        self.helpText = QtWidgets.QPlainTextEdit(HELP_TEXT)
        self.helpText.setWindowTitle("Help")
        self.helpText.setReadOnly(True)
        self.helpText.setMinimumSize(650, 350)

        central_widget = QtWidgets.QWidget()
        central_widget.setMinimumSize(700, 400)
        central_layout = QtWidgets.QVBoxLayout()
        central_widget.setLayout(central_layout)
        self.setCentralWidget(central_widget)

        self.param_group = QtWidgets.QGroupBox("Select your mailbox:")
        param_layout = QtWidgets.QVBoxLayout()

        self.mailbox_choice = QtWidgets.QComboBox()
        param_layout.addWidget(self.mailbox_choice)
        self.param_group.setLayout(param_layout)
        central_layout.addWidget(self.param_group, stretch=0)
        self.param_group.hide()

        self.find_button = QtWidgets.QPushButton("Search Duplicates")
        self.find_button.setFixedWidth(120)
        #self.find_button.setEnabled(False) #At startup, it is disabled
        central_layout.addWidget(self.find_button, stretch=0)

        self.progress_group = QtWidgets.QGroupBox("Progress:")
        progress_layout = QtWidgets.QHBoxLayout()
        self.progress_label = QtWidgets.QLabel("Progress of step 1/3:")
        self.progress_bar = QtWidgets.QProgressBar()
        self.cancel_button = QtWidgets.QPushButton("Cancel")
        progress_layout.addWidget(self.progress_label)
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.cancel_button)
        self.progress_group.setLayout(progress_layout)
        central_layout.addWidget(self.progress_group, stretch=0)

        self.result_group = QtWidgets.QGroupBox("Duplicates that can be deleted:")
        result_layout = QtWidgets.QVBoxLayout()
        self.selectAll_checkbox = QtWidgets.QCheckBox("Select All")
        self.selectAll_checkbox.setChecked(True)
        self.result_table = QtWidgets.QTableView()
        self.result_label = QtWidgets.QLabel("")
        result_buttons_layout = QtWidgets.QHBoxLayout()
        result_layout.addWidget(self.selectAll_checkbox, stretch = 0)
        result_layout.addWidget(self.result_table, stretch = 1) #If we stretch the main window, result_table will stretch
        result_layout.addWidget(self.result_label, stretch = 0)
        result_layout.addLayout(result_buttons_layout, stretch = 0)

        self.export_button = QtWidgets.QPushButton("Copy to Clipboard")
        self.export_button.setFixedWidth(120)
        self.delete_button = QtWidgets.QPushButton("Delete selected mails")
        self.delete_button.setFixedWidth(120)
        result_buttons_layout.addWidget(self.export_button, stretch = 0)
        result_buttons_layout.addWidget(self.delete_button, stretch = 0)

        self.result_group.setLayout(result_layout)
        central_layout.addWidget(self.result_group, stretch=1)

        self.result_group.hide() #will be displayed once we have results
        central_layout.addStretch() #allows window to stretch nicely while there are no results

class MainWindow(ui_MainWindow): #Adds logic to the interface widgets
    def __init__(self, FOLDER_SIZE_LIMIT, EXCLUDED_FOLDERS, INCLUDE_SUBFOLDERS):
        logging.info(f"Running Outlook Cleaner version {VERSION}")
        super().__init__()
        self.interrupt = False #Flag that will be triggered if user chooses to cancels a background job
        self.threadpool = QtCore.QThreadPool() #Pool for workers threads that will perform back ground work on emails

        self.help_action.triggered.connect(lambda x: self.helpText.show())
        self.mailbox_choice.currentIndexChanged.connect(self.mailbox_chosen_f)
        self.find_button.clicked.connect(self.launch_search_f)
        self.cancel_button.clicked.connect(lambda x: setattr(self, "interrupt", True)) #Signal to interrupt progress
        self.selectAll_checkbox.stateChanged.connect(self.selectAll_changed_f)
        self.export_button.clicked.connect(self.export_result_f)
        self.delete_button.clicked.connect(self.delete_selected_f)

        self.FOLDER_SIZE_LIMIT = FOLDER_SIZE_LIMIT
        self.EXCLUDED_FOLDERS = EXCLUDED_FOLDERS
        self.INCLUDE_SUBFOLDERS = INCLUDE_SUBFOLDERS
        self.del_model = None

        self.initiate_new_step("Retrieving mailboxes")
        worker_mailbox = Worker(self.list_mailboxes_b) #Create a thread to get the list of mailboxes
        worker_mailbox.signals.result.connect(self.list_mailboxes_b_result)
        worker_mailbox.signals.error.connect(self.list_mailboxes_b_error)
        self.threadpool.start(worker_mailbox) #Launch thread

    def mailbox_chosen_f(self): #Triggered by mailbox_choice
        self.result_group.hide() #Hide results if we choose a different mailbox
        self.CurrentAccountName = self.mailbox_choice.currentText()

    def initiate_new_step(self, text="Launching...", minimum=0, maximum=100):   #
        self.find_button.setEnabled(False)
        self.mailbox_choice.setEnabled(False)
        self.progress_group.show()
        self.progress_label.setText(text)
        self.progress_bar.setRange(minimum, maximum)
        self.setProgress(minimum)

    def setProgress(self, value):
        self.progress_bar.setValue(value)

    def get_outlook_dispatch(self, test_existing = False):
        """Connects to outlook for the first time. But also reconnects if connection was lost.
        Because nfortunately, win32com is quite unstable when used in multithread."""
        if test_existing: #If a connection was previously set, we first test it
            try:
                a = self.namespace.Folders[self.CurrentAccountName]
                logging.debug("Connection still OK")
                return #Connection is OK, we can skip the rest
            except pythoncom.com_error as e:
                logging.debug("Com_error. We try to reconnect")
        try: #First connection, or reconnection required
            pythoncom.CoInitialize() #Must sometimes be called when win32 is used with multithreadin
            self.outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application") #Early binding. Also generates Outlook constants
            self.EarlyBinding = True
        except:
            self.outlook = win32com.client.Dispatch("Outlook.Application") #Late binding. Does not generate constants.
            self.EarlyBinding = False
        logging.debug(f"Running in early binding? {self.EarlyBinding}")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def list_mailboxes_b(self, progress_callback): #List all mailboxes at startup
        #pythoncom.CoInitialize()
        self.get_outlook_dispatch(False)
        accounts = self.outlook.Session.Accounts #list all mailboxes available for user
        for i in range (1, accounts.Count+1): #Populate combo box with list of mailboxes
            account = accounts.Item(i)
            self.mailbox_choice.addItem(account.DisplayName)

    def list_mailboxes_b_result(self):
        self.param_group.show()
        self.progress_group.hide()
        self.find_button.setEnabled(True)
        self.mailbox_choice.setEnabled(True)

    def list_mailboxes_b_error(self): #If we fail to list mailboxes, deactivate find_button
        self.outlook, self.namespace = None, None
        self.param_group.show()
        self.mailbox_choice.addItem("No local mailbox available")

    @QtCore.Slot()
    def launch_search_f(self): #Triggered by find_button
        self.OtherMessage = ""
        self.worker_search = Worker(self.search_duplicates_b)
        self.worker_search.signals.progress.connect(self.search_duplicates_b_progress)
        self.worker_search.signals.newStep.connect(self.initiate_new_step)
        self.worker_search.signals.result.connect(self.search_duplicates_b_return)
        self.worker_search.signals.error.connect(self.search_duplicates_b_error)
        self.threadpool.start(self.worker_search)

    @QtCore.Slot()
    def update_total(self): #Triggered each time the list of selected messages changes. Also called at init
        self.del_model.count_deletable = len(self.del_model._df)
        self.del_model.count_selected = len(self.del_model._df[self.del_model._df.toDelete==True]) #Nb of messages selected for deletion
        self.del_model.size_selected = self.del_model._df[self.del_model._df.toDelete==True]["mess_size"].sum() #Cumulated size
        self.del_model.partially_checked = False
        if self.del_model.count_deletable == self.del_model.count_selected: #Adapt the selectAll checkbox status
            self.selectAll_checkbox.setCheckState(QtCore.Qt.Checked)
        elif self.del_model.count_selected == 0:
            self.selectAll_checkbox.setCheckState(QtCore.Qt.Unchecked)
        else:
            self.del_model.partially_checked = True
            self.selectAll_checkbox.setCheckState(QtCore.Qt.PartiallyChecked)
        result_text = f"{self.del_model.count_selected} messages marked for deletion out of {self.total_email_count} in inbox, representing {self.del_model.size_selected/1024/1000:.1f} Mo."
        self.result_label.setText(result_text + " " + self.OtherMessage)
        #self.OtherMessage = "" #TBD Clear message only when needed

    @QtCore.Slot()
    def selectAll_changed_f(self, status): #Triggered by the checkbox to select/unselect all messages
        self.del_model.beginResetModel() #Warn the model that data will change in background
        if status == QtCore.Qt.Checked:
            self.del_model._df.toDelete = True #Select all messages for deletion
        elif status == QtCore.Qt.Unchecked:
            self.del_model._df.toDelete = False #Deselect all messages
        elif status == QtCore.Qt.PartiallyChecked:
            if not(self.del_model.partially_checked): #Partially_checked can only be set by self.update_total, not by user
                self.selectAll_checkbox.setCheckState(QtCore.Qt.Checked)
        self.del_model.endResetModel()  #Inform data changed, and view should be updated

    def print_error_detail(self):
        ex_type, ex_value, ex_traceback = sys.exc_info() #details of the error
        trace_back = traceback.extract_tb(ex_traceback)
        stack_trace = []
        for trace in trace_back:
            #stack_trace.append(f"File: {trace[0]}, Line: {trace[1]}, Func.Name: {trace[2]}, Message: {trace[3]}")
            stack_trace.append(f"Line: {trace[1]}, Func.Name: {trace[2]}, Message: {trace[3]}")
        logging.debug(f"{ex_type} {ex_type.__name__} {ex_value}")
        logging.debug(stack_trace)

    def get_subFolders(self, top_folder, parent_folder='', n=1):
        """Recursively Returns all the sub-folders of top_folder if INCLUDE_SUBFOLDERS==True"""
        if top_folder.Name in self.EXCLUDED_FOLDERS:
            subFolders, email_count = [], 0
        else:
            email_count = top_folder.Items.Count
            if email_count < self.FOLDER_SIZE_LIMIT:
                subFolders = [(top_folder.Name, top_folder, parent_folder, email_count, n)]
            else:
                subFolders = []
                logging.info(f"{top_folder.Name} ignored for performance reasons because it has more than {self.FOLDER_SIZE_LIMIT} items")
            if self.INCLUDE_SUBFOLDERS:
                for f in  top_folder.Folders: #List sub-folders recursively
                    if f.DefaultMessageClass == "IPM.Note":
                        subFolders2, email_count2 = self.get_subFolders(f, top_folder, n+1)
                        subFolders += subFolders2
                        email_count += email_count2
        return subFolders, email_count

    def search_duplicates_b(self, progress_callback): #called in background by launch_search_f(self)
        self.worker_search.signals.newStep.emit("Counting emails...", 0, 100)
        self.get_outlook_dispatch(True)
        inbox = self.namespace.Folders[self.CurrentAccountName] #Inbox of the chosen mailbox
        #IDs of folders that are excluded by default. https://docs.microsoft.com/fr-fr/office/vba/api/outlook.oldefaultfolders
        if self.EarlyBinding:
            OTHER_EXCLUDED_FOLDERS_ID = [winConstants.olFolderSyncIssues, winConstants.olFolderContacts
                , winConstants.olFolderDrafts, winConstants.olFolderJournal, winConstants.olFolderRssFeeds]
        else:
            OTHER_EXCLUDED_FOLDERS_ID = [20, 10, 16, 11, 25]
        self.EXCLUDED_FOLDERS += [self.namespace.GetDefaultFolder(id).Name for id in OTHER_EXCLUDED_FOLDERS_ID]
        DELETED_FOLDER = self.namespace.GetDefaultFolder(winConstants.olFolderDeletedItems).Name #Deleted Items Folder name

        try:
            top_folder = inbox
            #top_folder = inbox.Folders["4-SUPP"].Folders["Test WIP"]
            #top_folder = inbox.Folders["Boîte de réception"] Éléments supprimés
            #top_folder = inbox.Folders["Éléments supprimés"]
        except:
            logging.critical("You have forced top_folder to a value that does not exist in this mailbox. Please modify code")
        self.StoreID = top_folder.StoreID
        self.cols_del = ['folderName', 'conversationID', 'ID', 'date', 'subject', 'topic', 'senderMail', 'toDelete', 'mess_size']
        folders, self.total_email_count = self.get_subFolders(top_folder, '', 1) #List all subfolders

        self.worker_search.signals.newStep.emit("Step 1/2 - Read all messages", 0, self.total_email_count)
        messageList = pd.DataFrame()
        deleteList = pd.DataFrame()
        count_read_mails = 0
        for folder in folders: #loop on folders to read all messages
            SubMessageList, SubDeleteList = self.build_messageList(folder[1], count_read_mails, progress_callback)
            messageList = pd.concat([messageList, SubMessageList], ignore_index=True)
            deleteList = pd.concat([deleteList, SubDeleteList], ignore_index=True)
            count_read_mails += folder[3]
        messageList = messageList.groupby('conversationID', sort=False)  #Group by conversation

        self.worker_search.signals.newStep.emit("Step 2/2 - Compare all messages", 0, len(messageList))
        count_topic_read = 0
        for conversationID, subList in messageList: #Loop on all conversation
            if self.interrupt:
                raise KeyboardInterrupt("User clicked on Cancel")
            for i, mess in subList.iterrows(): #Loop on all messages in a conversation
                date = mess.date
                subList2 = subList[ (subList['date']>date) & (subList['toDelete']==False) ] #List messages that are more recent:
                if mess.folderName != DELETED_FOLDER: #If the messsage is not Deleted, we don't want to compare it with Deleted messages
                    subList2 = subList2[subList2['folderName']!=DELETED_FOLDER ]
                for i2, mess2 in subList2.iterrows(): #Search messages in the same Conversation that contain mess
                    if all(att in mess2.atts for att in mess.atts): #First check: mess attachments are all in mess2
                        body_included = False
                        if mess.body in mess2.body: #Second check:mess body is fully included in mess2
                            body_included = True
                        elif "<mailto:" in mess.body or "<mailto:" in mess2.body: #Difference may com from formatted mail
                            if self.remove_mailtos(mess.body) in  self.remove_mailtos(mess2.body):
                                body_included = True
                        if body_included:
                            mess.body="" #We don't need body anymore
                            record = pd.DataFrame([[mess.folderName, mess.conversationID, mess.ID, mess.date, mess.subject, mess.topic, mess.senderName, mess.toDelete, mess.mess_size]], columns=self.cols_del)
                            record[["newerID", "newerFolder", "newerDate"]] = [mess2.ID, mess2.folderName, mess2.date] #Add info about the newer message
                            deleteList = pd.concat([deleteList, record], axis=0, ignore_index=True) #Add mess to the list of messages to delete
                            break #stop looping on subList2
            count_topic_read += 1
            self.worker_search.signals.progress.emit(count_topic_read) #Update progress bar
        return deleteList

    def search_duplicates_b_return(self, deleteList):
        self.del_model = Delete_TableModel(deleteList) #Populate table of results
        self.del_model.dataChanged.connect(self.update_total) #Receive signal sent by model when selected mails change
        self.del_model.modelReset.connect(self.update_total)  #Receive signal sent by model when df is updated after deletion
        self.update_total() #Initiate total
        self.progress_group.hide()
        self.find_button.setEnabled(True)
        self.mailbox_choice.setEnabled(True)
        self.result_group.show()    #Show widgets that will display results
        self.result_table.setModel(self.del_model) #Show results in TableView
        self.result_table.resizeColumnsToContents()

    def search_duplicates_b_progress(self, n,): #call back function to trak progress of search_duplicates_b()
        self.setProgress(n)

    def search_duplicates_b_error(self, err_tuple):
        self.progress_group.hide()
        self.find_button.setEnabled(True)
        self.mailbox_choice.setEnabled(True)
        #TBD: display a message indicating error
        self.interrupt = False

    def build_messageList(self, folder, count_read_mails, progress_callback):
        """Called by search_duplicates_b.
        Returns the list of messages in the folder, and deleteList containing undelivered and recalled mails"""
        messageList = []
        folderID = folder.EntryID
        folderName = folder.Name #folder.FullFolderPath
        messages = folder.Items
        for message in messages: #loop on all messages of folder
            if self.interrupt:
                raise KeyboardInterrupt("User clicked on Cancel")
            ID = message.EntryID
            date =  message.CreationTime #.date()   #message.Senton.date() does not always work
            date = datetime.datetime.combine(date.date(), date.time()) #convert to datetime
            subject = message.Subject
            topic = message.ConversationTopic.strip()
            conversationID = message.ConversationID
            isUnread = message.UnRead
            senderName = getattr(message, 'SenderName', None) #sometimes (messages undelivered or recalled) there is no sender
            senderMail = getattr(message, 'SenderEmailAddress', None)
            mess_size = message.Size
            atts = []
            if senderName is None: #Recalled or undelivered messages have no sender
                toDelete = isUnread==False #delete only unread messages
                body = ""
            else:
                toDelete = False
                body = message.Body #.HTMLbody, .RTFBody, .bodyFormat
                attachments = message.Attachments
                for att in attachments:
                    try:
                        attName = [getattr(att, 'FileName', att.DisplayName)]
                        atts += attName
                    except:
                        if att.Type != 6: #Normally, only type 6 should raise an error
                            logging.error(f"Type Error for: {date} | {subject} | {att.DisplayName} | {att.Type}")
                        #1-Office, 5-Mail, 6-Img (OLE Document), 7-Link Office
                        #https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.olattachmenttype?view=outlook-pia
                atts.sort()
            record = [[folderName, conversationID, ID, date, subject, topic, isUnread, senderName, senderMail, toDelete, body, atts, mess_size]]
            messageList += record
            count_read_mails += 1
            self.worker_search.signals.progress.emit(count_read_mails)
        cols = ["folderName", "conversationID", "ID", "date", "subject", "topic", "isUnread", "senderName", "senderMail", "toDelete", "body", "atts", "mess_size"]
        messageList = pd.DataFrame(messageList, columns = cols)
        deleteList = messageList[messageList['toDelete'] == True][self.cols_del]
        messageList = messageList[messageList['toDelete'] == False]
        messageList = messageList.sort_values(by=['date'], ascending=[True])
        return messageList, deleteList

    def remove_mailtos(self, body):
        """Called by search_duplicates_b. Remove the <mailto:> tag that is sometimes inserted after a mail"""
        pattern = " <mailto:[\w.\-]+@[\w.\-]+.\w+> " #search for "<mailto:" followed by a mail
        mailtos = re.findall(pattern, body)
        mailtos = list(dict.fromkeys(mailtos)) #remove duplicates
        for mailto in mailtos:
            body = body.replace(mailto, '')
        return body

    @QtCore.Slot()
    def export_result_f(self): #Triggered by export_button
        self.del_model._df[self.del_model.displayed_columns].to_clipboard() #Copy list of messages to clipboard

    @QtCore.Slot()
    def delete_selected_f(self): #Triggered by delete_button
        self.worker_delete = Worker(self.delete_selected_b) #Worker to delete in background
        self.initiate_new_step("Deleting...")
        self.worker_delete.signals.result.connect(self.delete_selected_b_result)
        self.worker_delete.signals.error.connect(self.delete_selected_b_error)
        self.worker_delete.signals.finished.connect(self.delete_selected_b_finished)
        self.threadpool.start(self.worker_delete)

    def delete_selected_b(self, progress_callback): #Called by delete_selected_f
        df_del = self.del_model._df[self.del_model._df.toDelete==True] #Only selected messages
        error_occured = False
        self.get_outlook_dispatch(True) #We test if connection to namespace is still active
        for i, mess in df_del.iterrows():
            try:
                outlook_mess = self.namespace.GetItemFromID(mess.ID, self.StoreID)
                outlook_mess.Delete() #Delete in Outlook
                logging.debug(f"Deleted: {mess.date} {mess.subject} {mess.folderName}")
            except BaseException as e: #We capture exception, so that next messages are handled
                error_occured = True
                error = e #We capture exception
                logging.error(f"Could not delete: {mess.date} {mess.subject} {mess.folderName}")
                self.print_error_detail()
                self.del_model._df.loc[self.del_model._df.ID==mess.ID, "toDelete"] = False #Acknowledge that mess was not deleted
                continue
        if error_occured:
            raise error #will trigger delete_selected_b_error, with the last error captured

    def delete_selected_b_result(self):
        self.OtherMessage = f"All messages successfully deleted"

    def delete_selected_b_error(self, error_tuple):
        self.OtherMessage = f"Error {error_tuple[0]}. Could not delete some emails"

    def delete_selected_b_finished(self):
        self.del_model.beginResetModel() #Notify the model that data will change in background
        self.del_model._df = self.del_model._df[self.del_model._df.toDelete==False] #Remove the deleted mails from the model
        self.del_model.endResetModel()  #Inform data changed, and view should be updated
        self.progress_group.hide() #Hide resul bar, and reactivate buttons
        self.find_button.setEnabled(True)
        self.mailbox_choice.setEnabled(True)

    def closeEvent(self, event): #Executed when we close the main window
        pythoncom.CoUninitialize()
        self.outlook, self.namespace = None, None #Clear references to win32com object. Does not work??
        #worker_quitOutlook = Worker(self.quitOutlook_b) #Clear references to win32com object. Does not work??
        #self.threadpool.start(worker_quitOutlook)
        self.helpText = None #Delete Help window
        self.threadpool = None
        event.accept()

    #def quitOutlook_b(self, progress_callback):
        #self.namespace.Logoff()
        #self.outlook.Application.Quit()   #Quit Outlook. Also quits outlook opened by the user !
        #self.outlook.Quit() #https://stackoverflow.com/questions/12705249/error-connecting-to-outlook-via-com

class Delete_TableModel(QtCore.QAbstractTableModel): #will hold the list of deleted mails to be displayed in QTableView

    def __init__(self, data):
        super().__init__()
        self._df = data
        self.displayed_columns = ["toDelete", "folderName", "subject", "date", "topic", "senderMail", "mess_size", "newerFolder", "newerDate"]
        self.displayed_header = ["Del", "Folder", "Subject", "Date", "Topic", "Sender", "Size (Mo)", "Newer mail in", "At date"]
        self.df_columns = list(self._df.columns)

    def rowCount(self, index):
        return self._df.shape[0]

    def columnCount(self, index):
        return len(self.displayed_columns)

    def headerData(self, section, orientation, role): #Return header of table
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                return self.displayed_header[section]
            if orientation == QtCore.Qt.Vertical:
                return None #str(self._df.index[section]) #Do not display line header

    def get_column_pos(self, index):
        column_name = self.displayed_columns[index.column()] #Columns are not displayed in the order of _df
        col = self.df_columns.index(column_name)
        return column_name, col

    def data(self, index, role): #Return data at index position
        if not index.isValid():
            return None
        column_name, col = self.get_column_pos(index)
        value = self._df.iloc[index.row(), col]
        if role == QtCore.Qt.DisplayRole: # or QtCore.Qt.EditRole:
            if column_name == "mess_size":
                return f"{value/1024/1000:.1f}" #Display in Mo with 1 decimal
            elif column_name == "toDelete":
                return None #We don't display a boolean, but a checkbox
            elif isinstance(value, datetime.datetime):
                try:
                    return value.strftime("%Y-%m-%d %H:%M:%S") #Format date
                except:
                    return value #Sometimes, we have an eror "NaTType does not support strftime"
            return str(value)
        elif role == QtCore.Qt.CheckStateRole:
            if column_name == "toDelete": #In this column, we display a checkbox
                return int(QtCore.Qt.Checked if value else QtCore.Qt.Unchecked)

    def flags(self, index): #flags to allow modifications
        column_name, col = self.get_column_pos(index)
        if column_name == "toDelete": #That column is UserCheckable
            return QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsEnabled|QtCore.Qt.ItemIsUserCheckable #Qt.ItemIsEditable
        return QtCore.Qt.ItemIsSelectable

    def setData(self, index, value, role): #Defines what to do when user changes data at index position
        if not index.isValid():
            return False
        if role == QtCore.Qt.CheckStateRole:
            column_name, col = self.get_column_pos(index)
            if column_name == "toDelete":
                checked = (value == int(QtCore.Qt.Checked))
                self._df.iloc[index.row(), col] = checked   #modify the value of toDelete in _df
                self.dataChanged.emit(index, index) #Emit a signal that data changed, and totals must be updated
                return True
        return False

def main():
    app = QtWidgets.QApplication(sys.argv) #Launch Qt
    myWindow = MainWindow(FOLDER_SIZE_LIMIT, EXCLUDED_FOLDERS, INCLUDE_SUBFOLDERS) #Create graphical interface
    myWindow.show() #display window
    sys.exit(app.exec()) #Launch Qt, and exit properly when window is closed

if __name__ == '__main__':
    logging.basicConfig(level=LOGGING_LEVEL, format="{asctime} {message}", style="{", datefmt="%H:%M:%S")
    main()

"""Technical Info:
N.B. It is possible to force the following parameters in the code:
    - limit the scan to a subfolder by changing the variable top_folder below in the code.
    - increase the FOLDER_SIZE_LIMIT that prevents scanning directories with too many mails. Beware of performance
    - force the exclusion of some subfolders in EXCLUDED_FOLDERS
1) Info on PySide6: https://www.pythonguis.com/
2) On Outlook APIs
    https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model
    https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem
    https://stackoverflow.com/questions/50127959/win32-dispatch-vs-win32-gencache-in-python-what-are-the-pros-and-cons
3) On Multithreading
    For cancel_button to work, we need background tasks in a separate thread
        https://www.pythonguis.com/tutorials/multithreading-pyside6-applications-qthreadpool/
    We made sure all win32com objects are in a unique threadpool. Otherwise, we would have to (painfully) apply:
        https://stackoverflow.com/questions/26764978/using-win32com-with-multithreading
Convention: functions launched by front are suffixed with _f, and running in back suffixed with _b

This script can be transformed in an executable via:
    python -m venv venv-outlook
    cd .\venv-outlook
    .\Scripts\activate
    python -m pip install pandas, pywin32, PySide6-Essentials, pyinstaller --trusted-host pypi.org --trusted-host files.pythonhosted.org
    pyinstaller.exe --hiddenimport win32timezone --onefile '.\Outlook - Cleaning.py' """