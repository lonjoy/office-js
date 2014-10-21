// ============================================================================
// office-js/text.js
// ============================================================================
// 
// Office365 client tests.
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

var office = require('./index');

// ============================================================================
// Main
// ============================================================================

var mailbox = new office.mail.Mailbox({
    user: 'foo@bar.baz',
    pass: 'thoushaltnot'
});

mailbox.setup(function (user, folders) {

    // read message
    folders[4].messages.read(function (messages) {
        console.log( messages[0] );
        console.log('read done');
    });

    // define message
    var message = {
        '@odata.type':'#Microsoft.OutlookServices.Message',
        'Subject'    :'Have you seen this new Mail REST API?',
        'Importance' :'High',
        'Body': {
            'ContentType': 'Text',
            'Content'    : 'It looks awesome!'
        },
        'ToRecipients': [{
          'Name':'Stelios Anagnostopoulos',
          'Address':'stean@microsoft.com'
        }]
    };

    // save message (api fails)
    mailbox.folders.drafts.messages.save(message, function (data) {
        console.log('save done');
    });

    // send message (api fails)
    mailbox.folders.drafts.messages.send(message, function (data) {
        console.log('send done');
    });
});
