// ============================================================================
// office-js/src/mail.js
// ============================================================================
// 
// Office365 mail client.
// 
// @authors 
//      10/20/2014 - stelios@outlook.com
//      
// All rights reserved.
// ============================================================================

'use strict';

// ============================================================================
// Imports
// ============================================================================

var xhr = require('./xhr');

// ============================================================================
// Cache
// ============================================================================

var Client = xhr.Client;

var serialize = JSON.stringify,
    parseJSON = JSON.parse;

// ============================================================================
// Mailbox
// ============================================================================

function Mailbox (config) {
    this.client   = new Client(config);
    this.folders  = new Folders(this);
    this.user     = new Person(this);
};

Mailbox.prototype = {

    createMessage: function (config) {
        return new Message(config, this.client);
    },

    setup: function (cb) {
        var that = this;
        this.user.read(function (user) {
            that.folders.read(function (folders) {
                cb( user, folders );
            });
        });
    }
};

// ============================================================================
// Person
// ============================================================================

function Person (mailbox) {
    this.mailbox = mailbox;
    this.client  = mailbox.client;
    this.path    = null;
    this.email   = null;
    this.name    = null;
    this.alias   = null;
};

Person.prototype = {

    get serialized () {
        return serialize({
            Name   : this.name,
            Address: this.email
        });
    },

    read: function (cb) {
        var that = this;
        this.client.get(this.path || '', function (data) {
            that.path  = data['@odata.id'];
            that.email = data.Id;
            that.name  = data.DisplayName;
            that.alias = data.Alias;
            cb( that ); 
        });
    }
};

// ============================================================================
// Message
// ============================================================================

function Message (data, messages) {
    this.messages = messages;
    this.client   = messages.client;
    if (!data['@odata.id']) {
        this.subject      = data.subject      || '';
        this.content      = data.content      || '';
        this.sender       = data.sender       || null;
        this.receivers    = data.receivers    || [];
        this.ccReceivers  = data.ccReceivers  || [];
        this.bccReceivers = data.bccReceivers || [];
        this.importance   = data.importance   || null;
        this.path         = null;
    } else {
        var mailbox = this.messages.folder.folders.mailbox;

        this.id     = data.Id;
        this.path   = data['@odata.id'];
        this.sender = new Person(data.Sender.EmailAddress, mailbox);

        this.receivers = data.ToRecipients.map(function (data) {
            return new Person( data.EmailAddress, mailbox );
        });

        this.ccReceivers = data.CcRecipients.map(function (data) {
            return new Person( data.EmailAddress, mailbox );
        });

        this.bccReceivers = data.BccRecipients.map(function (data) {
            return new Person( data.EmailAddress, mailbox );
        });

        this.dateCreated  = new Date(data.DateTimeCreated);
        this.dateReceived = new Date(data.DateTimeReceived);
        this.dateSent     = new Date(data.DateTimeSent);
        this.dateModified = new Date(data.DateTimeLastModified);

        this.importance                 = data.Importance;
        this.hasAttachments             = data.HasAttachments;
        this.isDeliveryReceiptRequested = data.IsDeliveryReceiptRequested;
        this.isReadReceiptRequested     = data.IsReadReceiptRequested;
        this.isDraft                    = data.IsDraft;
        this.isRead                     = data.IsRead;
        
        this.subject = data.Subject;
        this.preview = data.BodyPreview;
        this.content = data.Body.Content;//parseXML(data.Body.Content);
    }
};

Message.prototype = {

    get serialized () {
        var data = {};

        data.Sender  = this.sender.serialized;

        var idx = this.receivers.length;
        data.ToRecipients = new Array(idx);
        while (idx--)
            data.ToRecipients[idx] = {
                Name   : this.receivers[idx].name,
                Address: this.receivers[idx].email
            };

        data.Subject = this.subject;
        data.Body = {
            Content: this.content,
            ContentType: 'HTML'
        };

        if (typeof this.importance === 'boolean')
            data.Importance = this.importance ? 'High' : 'Low';
        else if (typeof this.importance === 'string')
            data.Importance = this.importance;

        data['@odata.type'] = '#Microsoft.OutlookServices.Message';

        return serialize(data);
    },

    // save: function (cb) {
    //     var that = this;
    //     this.client.post(this.path, this.serialized, function (data) {
    //         that.path      =  data['@odata.id'].replace('https://' + that.client.host, '');
    //         that.sender    = new Person( data.Sender.Address );
    //         that.receivers = data.ToRecipients.map(function (data) {
    //             return new Person( data.EmailAddress );
    //         });

    //         that.dateCreated  = new Date( data.DateTimeCreated );
    //         that.dateReceived = new Date( data.DateTimeReceived );
    //         that.dateSent     = new Date( data.DateTimeSent );
    //         that.dateModified = new Date( data.DateTimeLastModified );

    //         that.importance                 = data.Importance;
    //         that.hasAttachments             = data.HasAttachments;
    //         that.isDeliveryReceiptRequested = data.IsDeliveryReceiptRequested;
    //         that.isReadReceiptRequested     = data.IsReadReceiptRequested;
    //         that.isDraft                    = data.IsDraft;
    //         that.isRead                     = data.IsRead;
            
    //         that.subject = data.Subject;
    //         that.preview = data.BodyPreview;
    //         that.content = data.Body.Content;
    //         cb( that );
    //     });
    // }
};

// ============================================================================
// Message Collection
// ============================================================================

function Messages (path, folder) {
    this.path    = path;
    this.folder  = folder;
    this.client  = folder.client;
    this.list    = [];
    this.loaded  = false;
};

Messages.prototype = {

    send: function (message, cb) {
        var msg;
        if ( message instanceof Message || message.receivers )
            msg = message.serialized;
        else
            msg = serialize(message);
        var that = this;

        that.client.post('/sendmail', msg, function (data) {
            cb(data); // towait+do
        });
    },

    save: function (message, cb) {
        var msg;
        if ( message instanceof Message || message.receivers )
            msg = message.serialized;
        else
            msg = serialize(message);
        var that = this;

        that.client.post(this.path, msg, function (data) {
            cb(data); // towait+do
        });
    },

    read: function (cb) {
        if (this.loaded)
            return cb( this.list );

        var that = this;
        this.client.get(this.path, function (data) {

            if ( that.path !== data['@odata.nextLink'] ) {
                that.path = data['@odata.nextLink'];
                
                var messages = data.value || [];
                for (var i = 0, len = messages.length; i < len; i++)
                    messages[i] = new Message( messages[i], that );
                
                that.list = that.list.concat( messages );

                return cb( messages );
            
            } else {
                that.loaded = true;
                return cb( that.list );
            }
        });
    }
};

// ============================================================================
// Folder
// ============================================================================

function Folder (data, folders) {
    this.folders        = folders;
    this.client         = folders.client;
    this.id             = data.Id;
    this.path           = data['@odata.id'];
    this.name           = data.DisplayName;
    this.subfolderCount = data.ChildFolderCount;
    this.messages       = new Messages(this.path + '/messages', this);
};

Folder.prototype = {

    readAttributes: function (cb) {
        var that = this;
        this.client.get(this.path, function (data) {
            that.path           = data['@odata.id'];
            that.cpath          = data['ChildFolders@odata.navigationLink'];
            that.mpath          = data['Messages@odata.navigationLink'];
            that.name           = data.DisplayName;
            that.totalCount     = data.TotalCount;
            that.undreadCount   = data.UnreadItemCount;
            that.subfolderCount = data.ChildFolders;
            cb( that );
        });
    },

    readSubfolders: function (cb) {
    }
};

// ============================================================================
// Folder Collection
// ============================================================================

function Folders (mailbox) {
    this.path    = '/folders';
    this.mailbox = mailbox;
    this.client  = mailbox.client;
    this.list    = [];
    this.loaded  = false;
    this.drafts  = null
    this.deleted = null;
    this.inbox   = null;
};

Folders.prototype = {
    read: function (cb) {
        if (this.loaded)
            return cb( this.list );

        var that = this;
        this.client.get(this.path, function (data) {

            if ( that.path !== data['@odata.nextLink'] ) {
                that.path = data['@odata.nextLink'];
                var folders = data.value;
                for (var i = 0, len = folders.length; i < len; i++) {
                    folders[i] = new Folder( folders[i], that );
                    switch (folders[i].name) {
                        case 'Deleted Items':
                            that.deleted = folders[i];
                            break;
                        case 'Drafts':
                            that.drafts = folders[i];
                            break;
                        case 'Inbox':
                            that.inbox = folders[i];
                            break;
                        default:
                            break;
                    }
                }
                that.list = that.list.concat( folders );
                return cb( folders );
            } else {
                that.loaded = true;
                return cb( that.list );
            }
        });
    }
};

// ============================================================================
// Exports
// ============================================================================

module.exports = {
    Mailbox: Mailbox
};
