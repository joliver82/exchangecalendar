/* ***** BEGIN LICENSE BLOCK *****
 * Version: GPL 3.0
 *
 * The contents of this file are subject to the General Public License
 * 3.0 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.gnu.org/licenses/gpl.html
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * -- Exchange 2007/2010 Calendar and Tasks Provider.
 * -- For Thunderbird with the Lightning add-on.
 *
 * This work is a combination of the Storage calendar, part of the default Lightning add-on, and 
 * the "Exchange Data Provider for Lightning" add-on currently, october 2011, maintained by Simon Schubert.
 * Primarily made because the "Exchange Data Provider for Lightning" add-on is a continuation 
 * of old code and this one is build up from the ground. It still uses some parts from the 
 * "Exchange Data Provider for Lightning" project.
 *
 * Author: Michel Verbraak (info@1st-setup.nl)
 * Website: http://www.1st-setup.nl/wordpress/?page_id=133
 * email: exchangecalendar@extensions.1st-setup.nl
 *
 * Contributor: Krzysztof Nowicki (krissn@op.pl)
 * 
 *
 * This code uses parts of the Microsoft Exchange Calendar Provider code on which the
 * "Exchange Data Provider for Lightning" was based.
 * The Initial Developer of the Microsoft Exchange Calendar Provider Code is
 *   Andrea Bittau <a.bittau@cs.ucl.ac.uk>, University College London
 * Portions created by the Initial Developer are Copyright (C) 2009
 * the Initial Developer. All Rights Reserved.
 *
 * ***** BEGIN LICENSE BLOCK *****/

var Cc = Components.classes;
var Ci = Components.interfaces;

var Cr = Components.results;

ChromeUtils.import("resource://gre/modules/Services.jsm");
ChromeUtils.import("resource://gre/modules/AddonManager.jsm");

const { cal } = ChromeUtils.import("resource://calendar/modules/calUtils.jsm");

ChromeUtils.import("resource://exchangecommon/ecFunctions.js");
ChromeUtils.import("resource://exchangecommon/soapFunctions.js");
ChromeUtils.import("resource://exchangecommon/ecExchangeRequest.js");

ChromeUtils.import("resource://exchangecommon/erFindFolder.js");
ChromeUtils.import("resource://exchangecommon/erGetFolder.js");

ChromeUtils.import("resource://exchangecommon/erGetItems.js");
ChromeUtils.import("resource://exchangecommon/erCreateItem.js");
ChromeUtils.import("resource://exchangecommon/erUpdateItem.js");
ChromeUtils.import("resource://exchangecommon/erDeleteItem.js");

ChromeUtils.import("resource://exchangecommon/erSyncFolderItems.js");
ChromeUtils.import("resource://exchangecommon/erGetUserAvailability.js");

ChromeUtils.import("resource://exchangecalendar/erFindCalendarItems.js");
ChromeUtils.import("resource://exchangecalendar/erFindTaskItems.js");
ChromeUtils.import("resource://exchangecalendar/erFindFollowupItems.js");

ChromeUtils.import("resource://exchangecalendar/erFindMasterOccurrences.js");
ChromeUtils.import("resource://exchangecalendar/erGetMasterOccurrenceId.js");
ChromeUtils.import("resource://exchangecalendar/erGetMeetingRequestByUID.js");
ChromeUtils.import("resource://exchangecalendar/erFindOccurrences.js");
ChromeUtils.import("resource://exchangecalendar/erGetOccurrenceIndex.js");

ChromeUtils.import("resource://exchangecalendar/erSendMeetingRespons.js");
ChromeUtils.import("resource://exchangecalendar/erSyncInbox.js");

ChromeUtils.import("resource://exchangecalendar/erCreateAttachment.js");
ChromeUtils.import("resource://exchangecalendar/erDeleteAttachment.js");

ChromeUtils.import("resource://exchangecommoninterfaces/xml2json/xml2json.js");

ChromeUtils.import("resource://interfacescalendartask/exchangeTodo/mivExchangeTodo.js");
ChromeUtils.import("resource://interfacescalendartask/exchangeEvent/mivExchangeEvent.js");

var EXPORTED_SYMBOLS = ["calExchangeTest"];

var globalStart = new Date().getTime();

const nsIAP = Ci.nsIActivityProcess;
const nsIAE = Ci.nsIActivityEvent;
const nsIAM = Ci.nsIActivityManager;

var gActivityManager;

if (Cc["@mozilla.org/activity-manager;1"]) {
    gActivityManager = Cc["@mozilla.org/activity-manager;1"].getService(nsIAM);
    (new (ChromeUtils.import("resource://exchangecommoninterfaces/global/mivFunctions.js").mivFunctions)()).LOG("-- ActivityManager available. Enabling it.");
}
else {
    (new (ChromeUtils.import("resource://exchangecommoninterfaces/global/mivFunctions.js").mivFunctions)()).LOG("-- ActivityManager not available.");
}


const fieldPathMap = {
    'ActualWork': 'task',
    'AdjacentMeetingCount': 'calendar',
    'AdjacentMeetings': 'calendar',
    'AllowNewTimeProposal': 'calendar',
    'AppointmentReplyTime': 'calendar',
    'AppointmentSequenceNumber': 'calendar',
    'AppointmentState': 'calendar',
    'AssignedTime': 'task',
    'AssociatedCalendarItemId': 'meeting',
    'Attachments': 'item',
    'BillingInformation': 'task',
    'Body': 'item',
    'CalendarItemType': 'calendar',
    'Categories': 'item',
    'ChangeCount': 'task',
    'Companies': 'task',
    'CompleteDate': 'task',
    'ConferenceType': 'calendar',
    'ConflictingMeetingCount': 'calendar',
    'ConflictingMeetings': 'calendar',
    'Contacts': 'task',
    'ConversationId': 'item',
    'Culture': 'item',
    'DateTimeCreated': 'item',
    'DateTimeReceived': 'item',
    'DateTimeSent': 'item',
    'DateTimeStamp': 'calendar',
    'DelegationState': 'task',
    'Delegator': 'task',
    'DeletedOccurrences': 'calendar',
    'DisplayCc': 'item',
    'DisplayTo': 'item',
    'DueDate': 'task',
    'Duration': 'calendar',
    'EffectiveRights': 'item',
    'End': 'calendar',
    'EndTimeZone': 'calendar',
    'FirstOccurrence': 'calendar',
    'FolderClass': 'folder',
    'FolderId': 'folder',
    'HasAttachments': 'item',
    'HasBeenProcessed': 'meeting',
    'Importance': 'item',
    'InReplyTo': 'item',
    'IntendedFreeBusyStatus': 'meetingRequest',
    'InternetMessageHeaders': 'item',
    'IsAllDayEvent': 'calendar',
    'IsAssignmentEditable': 'task',
    'IsAssociated': 'item',
    'IsCancelled': 'calendar',
    'IsComplete': 'task',
    'IsDelegated': 'meeting',
    'IsDraft': 'item',
    'IsFromMe': 'item',
    'IsMeeting': 'calendar',
    'IsOnlineMeeting': 'calendar',
    'IsOutOfDate': 'meeting',
    'IsRecurring': 'calendar',
    'IsResend': 'item',
    'IsResponseRequested': 'calendar',
    'IsSubmitted': 'item',
    'IsTeamTask': 'task',
    'IsUnmodified': 'item',
    'ItemClass': 'item',
    'messageId': 'item',
    'ItemId': 'item',
    'LastModifiedName': 'item',
    'LastModifiedTime': 'item',
    'LastOccurrence': 'calendar',
    'LegacyFreeBusyStatus': 'calendar',
    'Location': 'calendar',
    'MeetingRequestType': 'meetingRequest',
    'MeetingRequestWasSent': 'calendar',
    'MeetingTimeZone': 'calendar',
    'MeetingWorkspaceUrl': 'calendar',
    'Mileage': 'task',
    'MimeContent': 'item',
    'ModifiedOccurrences': 'calendar',
    'MyResponseType': 'calendar',
    'NetShowUrl': 'calendar',
    'OptionalAttendees': 'calendar',
    'Organizer': 'calendar',
    'OriginalStart': 'calendar',
    'Owner': 'task',
    'ParentFolderId': 'item',
    'PercentComplete': 'task',
    'Recurrence': 'calendar',
    'RecurrenceId': 'calendar',
    'ReminderDueBy': 'item',
    'ReminderIsSet': 'item',
    'ReminderMinutesBeforeStart': 'item',
    'RequiredAttendees': 'calendar',
    'Resources': 'calendar',
    'ResponseObjects': 'item',
    'ResponseType': 'meeting',
    'SearchParameters': 'folder',
    'Sensitivity': 'item',
    'Size': 'item',
    'Start': 'calendar',
    'StartDate': 'task',
    'StartTimeZone': 'calendar',
    'StatusDescription': 'task',
    'Status': 'task',
    'Subject': 'item',
    'TimeZone': 'calendar',
    'TotalWork': 'task',
    'UID': 'calendar',
    'UniqueBody': 'item',
    'WebClientEditFormQueryString': 'item',
    'WebClientReadFormQueryString': 'item',
    'When': 'calendar'
};

const dayRevMap = {
    'MO': 'Monday',
    'TU': 'Tuesday',
    'WE': 'Wednesday',
    'TH': 'Thursday',
    'FR': 'Friday',
    'SA': 'Saturday',
    'SU': 'Sunday'
};

const dayIdxMap = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

const weekRevMap = {
    '1': 'First',
    '2': 'Second',
    '3': 'Third',
    '4': 'Fourth',
    '-1': 'Last'
};

const monthIdxMap = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];

const participationMap = {
    "Unknown": "NEEDS-ACTION",
    "NoResponseReceived": "NEEDS-ACTION",
    "Tentative": "TENTATIVE",
    "Accept": "ACCEPTED",
    "Decline": "DECLINED",
    "Organizer": "ACCEPTED"
};

const dayMap = {
    'Monday': 'MO',
    'Tuesday': 'TU',
    'Wednesday': 'WE',
    'Thursday': 'TH',
    'Friday': 'FR',
    'Saturday': 'SA',
    'Sunday': 'SU',
    'Weekday': ['MO', 'TU', 'WE', 'TH', 'FR'],
    'WeekendDay': ['SA', 'SO'],
    'Day': ['MO', 'TU', 'WE', 'TH', 'FR', 'SA', 'SO']
};

const weekMap = {
    'First': 1,
    'Second': 2,
    'Third': 3,
    'Fourth': 4,
    'Last': -1
};

const monthMap = {
    'January': 1,
    'February': 2,
    'March': 3,
    'April': 4,
    'May': 5,
    'June': 6,
    'July': 7,
    'August': 8,
    'September': 9,
    'October': 10,
    'November': 11,
    'December': 12
};

const MAPI_PidLidTaskAccepted = "33032";
const MAPI_PidLidTaskLastUpdate = "33045";
const MAPI_PidLidTaskHistory = "33050";
const MAPI_PidLidTaskOwnership = "33065";
const MAPI_PidLidTaskMode = "34072";
const MAPI_PidLidTaskGlobalId = "34073";
const MAPI_PidLidTaskAcceptanceState = "33066";
const MAPI_PidLidReminderSignalTime = "34144";
const MAPI_PidLidReminderSet = "34051";

/**
 * calExchangeTest
 * 
 * @class
 * @constructor
 */
function calExchangeTest() {

    try {
        this.myId = null;

        this.initProviderBase();

        this.globalFunctions = (new (ChromeUtils.import("resource://exchangecommoninterfaces/global/mivFunctions.js").mivFunctions)());

        this.globalFunctions.LOG("Constructor: 1");

        this.timeZones = (new (ChromeUtils.import("resource://interfacescalendartask/exchangeTimeZones/mivExchangeTimeZones.js").mivExchangeTimeZones)());

        this.noDB = true;
        this.dbInit = false;

        this.folderPathStatus = 1;
        this.firstrun = true;
        this.mUri = "";
        this.mid = null;

        //	this.initialized = false;

        //	this.prefs = null;
        this.mUseOfflineCache = null;
        this.mNotConnected = true;

        this.myAvailable = false;

        this.mPrefs = null;

        this.itemCacheById = {};
        this.itemCancelQueue = {};

        this.itemCacheByStartDate = {};
        this.itemCacheByEndDate = {};
        this.recurringMasterCache = {};
        this.recurringMasterCacheById = {};
        this.newMasters = {};
        this.parentLessItems = {};

        this.startDate = null;
        this.endDate = null;

        this.syncState = null;
        this.syncStateInbox = null;
        this.syncInboxState = null;
        this._weAreSyncing = false;
        this.firstSyncDone = false;

        this.meetingRequestsCache = [];
        this.meetingCancelationsCache = [];
        this.meetingrequestAnswered = [];
        this.meetingResponsesCache = [];

        this.getItemSyncQueue = [];
        this.getItemsSyncQueue = [];
        this.processItemSyncQueueBusy = false;

        this.offlineTimer = null;
        this.offlineQueue = [];

        this.doReset = false;

        this.shutdown = false;

        this.globalFunctions.LOG("Constructor: 2");

        this.inboxPoller = Cc["@mozilla.org/timer;1"]
            .createInstance(Ci.nsITimer);

        this.cacheLoader = Cc["@mozilla.org/timer;1"]
            .createInstance(Ci.nsITimer);
        this.loadingFromCache = false;

        this.observerService = Cc["@mozilla.org/observer-service;1"]
            .getService(Ci.nsIObserverService);

        this.globalFunctions.LOG("Constructor: 3");

        this.lightningNotifier = (new (ChromeUtils.import("resource://exchangecommoninterfaces/exchangeLightningNotifier/mivExchangeLightningNotifier.js").mivExchangeLightningNotifier)());

        this.loadBalancer = (new (ChromeUtils.import("resource://exchangecommoninterfaces/exchangeLoadBalancer/mivExchangeLoadBalancer.js").mivExchangeLoadBalancer)());

        this.exchangeStatistics = (new (ChromeUtils.import("resource://exchangecommoninterfaces/exchangeStatistics/mivExchangeStatistics.js").mivExchangeStatistics)());

        this.globalFunctions.LOG("Constructor: 4");

        this.calendarPoller = null;

        this.mObserver = new ecObserver(this);

        this.supportsTasks = false;
        this.supportsEvents = false;

        this.folderProperties = null;
        this._readOnly = true;
        this.folderIsNotAvailable = true;

        this.exporting = false;
        this.OnlyShowAvailability = false;

        this.updateCalendarItems = [];
        this.updateCalendarTimer = Cc["@mozilla.org/timer;1"]
            .createInstance(Ci.nsITimer);
        this.updateCalendarTimerRunning = false;

        this._canDelete = false;
        this._canModify = false;
        this._canCreateContent = false;

        this.globalFunctions.LOG("Constructor: 5");

        this.mIsOffline = Components.classes["@mozilla.org/network/io-service;1"]
            .getService(Components.interfaces.nsIIOService).offline;

        this._exchangeCurrentStatus = Cr.NS_OK; //Cr.NS_ERROR_FAILURE; //Cr.NS_OK;

        this._connectionStateDescription = "";
        //this.globalFunctions.LOG("Our offline status is:"+this.mIsOffline+".");

        this.itemCount = 0;
        this.itemUpdates = 0;
        this.itemsFromExchange = 0;
        this.masterCount = 0;

        this.globalFunctions.LOG("Constructor: 6");
        
    }
    catch (err) {
        dump("mivExchangeCalendar.new Err:" + err + "\n");
        this.globalFunctions.LOG("Constructor: mivExchangeCalendar.new Err:" + err);
    }

    exchWebService.check4addon.logAddOnVersion();
}

var calExchangeTestGUID = "518419a0-1997-11eb-8b6f-0800200c9a66";
var calExchangeTestClassID = Components.ID("{" + calExchangeTestGUID + "}");

// limited interfaces to minimum
var calExchangeTestInterfaces = [
    Ci.calICalendar,
    Ci.calISchedulingSupport
];
    
var calExchangeTestDescription = "Exchange 2007/2010 Calendar and Tasks Provider";
var calExchangeTestContractID = "@mozilla.org/calendar/calendar;1?type=exchangecalendar";

calExchangeTest.prototype = {

    get timeStamp() {
        var elapsed = new Date().getTime() - globalStart;
        //dump("elapsed:"+elapsed);
        return elapsed;
    },

    __proto__: cal.provider.BaseClass.prototype,

    // Begin nsIClassInfo
    classID: calExchangeTestClassID,
    contractID: calExchangeTestContractID,
    classDescription: calExchangeTestDescription,
    
    QueryInterface: cal.generateQI(calExchangeTestInterfaces),
    classInfo: cal.generateCI({
        classDescription: calExchangeTestDescription,
        contractID: calExchangeTestContractID,
        classID: calExchangeTestClassID,
        interfaces: calExchangeTestInterfaces
    }),
    
    // Begin calICalendar

    //  attribute AUTF8String id;

    //   attribute AUTF8String name;

    //  readonly attribute AUTF8String type;
    get type() {
        return "exchangecalendar";
    },

    //  readonly attribute AString providerID;
    get providerID() {
        return "exchangecalendar@extensions.1st-setup.nl";
    },

    //  attribute calICalendar superCalendar;

    get id() {
        return this.myId;
    },

    set id(aValue) {
        // We ignore this.
        //dump("Someone is setting the id to '"+aValue+"' for calendar:"+this.name+"\n");
    },

    //  attribute nsIURI uri;
    get uri() {
        return this.mUri;
    },

    set uri(aUri) {
        this.myId = aUri.pathQueryRef.substr(1);
        this.mUri = aUri;

        return this.mUri;
    },

    set readOnly(aValue) {
        //dump("set readOnly:"+this.name+"|"+this.globalFunctions.STACK(10)+"\n");
        this.prefs.setBoolPref("UserReadOnly", aValue);
        this.readOnlyInternal = aValue;
    },

    get readOnly() {
        var userPref = this.globalFunctions.safeGetBoolPref(this.prefs, "UserReadOnly", false);
        if (userPref === true) return true;
        return this.readOnlyInternal;
    },

    set readOnlyInternal(aValue) {
        //dump("set readOnlyInternal:"+this.name+"\n");

        this._readOnly = aValue;
    },

    //  attribute boolean readOnly;
    get readOnlyInternal() {
        //dump("get readOnly: name:"+this.name+", this._readOnly:"+this._readOnly+", this.notConnected:"+this.notConnected+"\n");
        return (this._readOnly);
    },

    //  attribute boolean transientProperties;

    //  nsIVariant getProperty(in AUTF8String aName);
    getProperty: function _getProperty(aName) {
        //	if (!this.isInitialized) {
        //		return;
        //	}

        //dump("2 getProperty("+aName+")\n");
        switch (aName) {
        case "exchWebService.offlineCacheDBHandle":
            return this.offlineCacheDB;

        case "exchWebService.offlineOrNotConnected":
            return (this.isOffline);

        case "exchWebService.useOfflineCache":
            return this.useOfflineCache;
        case "exchWebService.getFolderProperties":
            this.globalFunctions.LOG("Requesting exchWebService.getFolderProperties property.");
            if (this.folderProperties) {
                return this.folderProperties.toString();
            }
            return null;
            break;
        case "exchWebService.checkFolderPath":
            this.globalFunctions.LOG("Requesting exchWebService.checkFolderPath property.");
            this.checkFolderPath();
            return "ok";
            break;
        case "capabilities.tasks.supported":
            return this.supportsTasks;
            break;
        case "capabilities.events.supported":
            return this.supportsEvents;
            break;
        case "auto-enabled":
            return true;
        case "organizerId":
            return "mailto:" + this.mailbox;
            break;
        case "organizerCN":
            return this.userDisplayName;
            break;
        case "cache.supported":
            return false;
        case "requiresNetwork":
            return false;
        case "disabled":
            if (this.prefs) {
                this._disabled = this.globalFunctions.safeGetBoolPref(this.prefs, "disabled", false);
                if (this._disabled) return this._disabled;
            }

            return ((!this.isInitialized) && (this.folderPathStatus == 0));
        case "itip.notify-replies":
            return true;
        case "itip.transport":
            this.logInfo("getProperty: itip.transport");
            return this.QueryInterface(Ci.calIItipTransport);
            break;
            //return true;
        case "capabilities.autoschedule.supported":
            this.logInfo("capabilities.autoschedule.supported");
            return true;
        case "exchangeCurrentStatus":
            return this._exchangeCurrentStatus;
        }
        // itip.disableRevisionChecks

        // capabilities.events.supported
        // capabilities.tasks.supported

        //dump("1 getProperty("+aName+")="+this.__proto__.__proto__.getProperty.apply(this, arguments)+"\n");
        return this.__proto__.__proto__.getProperty.apply(this, arguments);
    },

    //  void setProperty(in AUTF8String aName, in nsIVariant aValue);
    setProperty: function setProperty(aName, aValue) {

        this.logInfo("setProperty. aName:" + aName + ", aValue:" + aValue);
        switch (aName) {
        case "exchangeCurrentStatus":
            //dump("name1:"+this.name+", exchangeCurrentStatus:"+this._exchangeCurrentStatus+", newStatus:"+aValue+"\n");
            var oldStatus = this._exchangeCurrentStatus;
            this._exchangeCurrentStatus = aValue;
            if (aValue != oldStatus) {
                //dump("name2:"+this.name+", exchangeCurrentStatus:"+aValue+"\n");
                this.observers.notify("onPropertyChanged", [this.superCalendar, "exchangeCurrentStatus", aValue, oldStatus]);
            }
            return;
        case "disabled":
            var oldDisabledState = this._disabled;
            this._disabled = aValue;
            this.prefs.setBoolPref("disabled", aValue);
            if ((aValue) && (oldDisabledState != this._disabled)) {
                //dump("Calendar is set to disabled\n");
                this.resetCalendar();
            }
            if ((!this._disabled) && (oldDisabledState != this._disabled)) {
                this.doReset = true;
                this.resetCalendar();
            }
            return;
        case "exchWebService.useOfflineCache":

            this.useOfflineCache = aValue;
            this.logInfo("setProperty: useOfflineCache = " + this.useOfflineCache + "  offlineCacheDB  " + this.offlineCacheDB);

            if (!aValue) {
                if (this.offlineCacheDB) {
                    try {
                        if (this.offlineCacheDB) this.offlineCacheDB.close();
                        this.offlineCacheDB = null;
                    }
                    catch (exc) {}
                }

                // Remove the offline cache database when we delete the calendar.
                if (this.dbFile) {
                    this.dbFile.remove(true);
                    this.offlineCacheDB = null;
                }
            }
            return;
        }

        this.__proto__.__proto__.setProperty.apply(this, arguments);

    },

    //  void deleteProperty(in AUTF8String aName);

    //  void addObserver( in calIObserver observer );
    //  void removeObserver( in calIObserver observer );

    //  calIOperation addItem(in calIItemBase aItem,
    //                in calIOperationListener aListener);
    addItem: function _addItem(aItem, aListener) {
        this.logInfo("addItem id=" + aItem.id + ", aItem.calendar:" + aItem.calendar);

        return this.adoptItem(newItem, aListener);
    },


    //  calIOperation adoptItem(in calIItemBase aItem,
    //                  in calIOperationListener aListener);
    adoptItem: function _adoptItem(aItem, aListener) {
        this.logInfo("adoptItem()");

        return;
    },


    //  calIOperation modifyItem(in calIItemBase aNewItem,
    //                   in calIItemBase aOldItem,
    //                   in calIOperationListener aListener);

    modifyItem: function _modifyItem(aNewItem, aOldItem, aListener) {

        this.logInfo("modifyItem");

        return null;
    },


    //  calIOperation deleteItem(in calIItemBase aItem,
    //                   in calIOperationListener aListener);
    deleteItem: function _deleteItem(aItem, aListener) {
        this.logInfo("deleteItem");

        return;
    },

    //  calIOperation getItem(in string aId, in calIOperationListener aListener);
    getItem: function _getItem(aId, aListener, aRetry) {
        this.logInfo("getItem: aId:" + aId);

        return;
    },


    //  calIOperation getItems(in unsigned long aItemFilter,
    //                 in unsigned long aCount,
    //                 in calIDateTime aRangeStart,
    //                 in calIDateTime aRangeEndEx,
    //                 in calIOperationListener aListener);
    getItems: function _getItems(aItemFilter, aCount,
        aRangeStart, aRangeEnd, aListener) {

        this.logInfo("getItems: aItemFilter, " + aItemFilter
            + " aCount, " + aCount
            + " aListener , " + aListener);


        return;
    },

    //  calIOperation refresh();
    refresh: function _refresh() {

        return;
    },

    // End calICalendar

    // Begin calISchedulingSupport
    //  boolean isInvitation(in calIItemBase aItem);
    isInvitation: function _isInvitation(aItem, ignoreStatus) {

        return false;
    },

    // boolean canNotify(in AUTF8String aMethod, in calIItemBase aItem);
    canNotify: function _canNotify(aMethod, aItem) {
        this.logInfo("canNotify: aMethod=" + aMethod + ":" + aItem.title);

        return true;
    },

    // calIAttendee getInvitedAttendee(in calIItemBase aItem);
    getInvitedAttendee: function _getInvitedAttendee(aItem) {
        return;
    },
    // End calISchedulingSupport



    /**
     * Internal logging function that should be called on any database error,
     * it will log as much info as possible about the database context and
     * last statement so the problem can be investigated more easilly.
     *
     * @param message           Error message to log.
     * @param exception         Exception that caused the error.
     */
    logError: function _logError(message, exception) {
        let logMessage = "(" + this.name + ") " + message;

        if (exception) {
            logMessage += "\nException: " + exception;
        }

        this.globalFunctions.ERROR(logMessage + "\n" + this.globalFunctions.STACK(10));
    },

    logInfo: function _logInfo(message) {
        this.globalFunctions.LOG("[" + this.name + "] " + message + " (" + this.globalFunctions.STACKshort() + ")");
    },

    logDebug: function _logDebug(message) {
        this.globalFunctions.DEBUG("[" + this.name + "] " + message + " (" + this.globalFunctions.STACKshort() + ")");
    },

};

function ecObserver(inCalendar) {
    this.calendar = inCalendar;

    var self = this;
    this.ecInvitationsCalendarManagerObserver = {
        onCalendarRegistered: function cMO_onCalendarRegistered(aCalendar) {
        },

        onCalendarUnregistering: function cMO_onCalendarUnregistering(aCalendar) {
            self.calendar.logInfo("onCalendarUnregistering name=" + aCalendar.name + ", id=" + aCalendar.id);
            if (aCalendar.id == self.calendar.id) {

                self.calendar.doDeleteCalendar();
                self.calendar.logInfo("Removing calendar preference settings.");

                var rmPrefs = Cc["@mozilla.org/preferences-service;1"]
                    .getService(Ci.nsIPrefService)
                    .getBranch("extensions.exchangecalendar@extensions.1st-setup.nl.");
                try {
                    rmPrefs.deleteBranch(aCalendar.id);
                }
                catch (err) {}

                aCalendar.removeFile("syncState.txt");
                aCalendar.removeFile("syncInboxState.txt");
                aCalendar.removeFile("folderProperties.txt");
                aCalendar.removeFile("syncStateInbox.txt");
                self.unregister();
            }
        },

        onCalendarDeleting: function cMO_onCalendarDeleting(aCalendar) {
            self.calendar.logInfo("onCalendarDeleting name=" + aCalendar.name + ", id=" + aCalendar.id);

        }
    };


    this.register();
}

ecObserver.prototype = {

    observe: function (subject, topic, data) {
        // Do your stuff here.
        //LOG("ecObserver.observe. topic="+topic+",data="+data+"\n"); 
        switch (topic) {
        case "onCalReset":
            if (data == this.calendar.id) {
                this.calendar.resetCalendar();
            }
            break;
        case "onExchangeConnectionError":
            var parts = data.split("|");
            if (this.calendar.serverUrl == parts[0]) {
                this.calendar.connectionIsNotOk(parts[2], parts[1]);
            }
            break;
        case "onExchangeConnectionOk":
            // See if it is for us
            if (data == this.calendar.serverUrl) {
                this.calendar.connectionIsOk();
            }
            break;
        case "quit-application":
            this.unregister();
            break;
        case "nsPref:changed":
            if ((data == "extensions.1st-setup.debug.log") || (data == "extensions.1st-setup.core.debuglevel")) {
                this.calendar.updateDoDebug();
            }
            break;
        case "network:offline-status-changed":
            this.calendar.offlineStateChanged(data);
            break;
        }
    },

    register: function () {
        var observerService = Cc["@mozilla.org/observer-service;1"]
            .getService(Ci.nsIObserverService);
        observerService.addObserver(this, "onCalReset", false);
        observerService.addObserver(this, "onExchangeConnectionError", false);
        observerService.addObserver(this, "onExchangeConnectionOk", false);
        observerService.addObserver(this, "quit-application", false);
        observerService.addObserver(this, "network:offline-status-changed", false);

        Services.prefs.addObserver("extensions.1st-setup.debug.log", this, false);
        Services.prefs.addObserver("extensions.1st-setup.core.debuglevel", this, false);


        cal.getCalendarManager().addObserver(this.ecInvitationsCalendarManagerObserver);
    },

    unregister: function () {
        this.calendar.doShutdown();

        var observerService = Cc["@mozilla.org/observer-service;1"]
            .getService(Ci.nsIObserverService);
        observerService.removeObserver(this, "onCalReset");
        observerService.removeObserver(this, "onExchangeConnectionError");
        observerService.removeObserver(this, "onExchangeConnectionOk");
        observerService.removeObserver(this, "quit-application");
        observerService.removeObserver(this, "network:offline-status-changed");

        cal.getCalendarManager().removeObserver(this.ecInvitationsCalendarManagerObserver);
    }
}



if (!exchWebService) var exchWebService = {};

exchWebService.check4addon = {

    alreadyLogged: false,

    checkAddOnIsInstalledCallback: function _checkAddOnIsInstalledCallback(aAddOn) {
        let mivFunctions = (new (ChromeUtils.import("resource://exchangecommoninterfaces/global/mivFunctions.js").mivFunctions)());
        if (!aAddOn) {
            mivFunctions.LOG("Exchange Calendar and Tasks add-on is NOT installed.");
        }
        else {
            mivFunctions.LOG(aAddOn.name + " is installed.");
            try {
                mivFunctions.LOG(aAddOn.name + " is installed from:" + aAddOn.sourceURI.prePath + aAddOn.sourceURI.pathQueryRef);
            }
            catch (err) {
                mivFunctions.LOG(aAddOn.name + " unable to determine where installed from.");
            }
            mivFunctions.LOG(aAddOn.name + " is version:" + aAddOn.version);
            if (aAddOn.isActive) {
                mivFunctions.LOG(aAddOn.name + " is active.");
            }
            else {
                mivFunctions.LOG(aAddOn.name + " is NOT active.");
            }
        }

    },

    logAddOnVersion: function _logAddOnVersion() {
        if (this.alreadyLogged) return;

        this.alreadyLogged = true;
        
        try{
            AddonManager.getAddonByID("exchangecalendar@extensions.1st-setup.nl", exchWebService.check4addon.checkAddOnIsInstalledCallback);
        }
        catch(ex)
        {
            var globalFunctions = (new (ChromeUtils.import("resource://exchangecommoninterfaces/global/mivFunctions.js").mivFunctions)());
            globalFunctions.LOG("logAddOnVersion:  Exception occured ");
        }
    }
};
