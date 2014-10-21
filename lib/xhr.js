// ============================================================================
// office-js/src/client.js
// ============================================================================
// 
// Https client for Office365.
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

var https = require('https');

// ============================================================================
// Cache
// ============================================================================

var request   = https.request,
    serialize = JSON.stringify,
    parseJSON = JSON.parse;

// ============================================================================
// Definition
// ============================================================================

function Client (config) {
    this.host = config.host || Client.config.host;
    this.port = config.port || Client.config.port;
    this.apiv = config.apiv || Client.config.apiv;
    this.user = config.user || Client.config.user;
    this.pass = config.pass || Client.config.pass;
    this.path = '/api/' + this.apiv + '/me';
};

Client.config = {
    host: 'outlook.office365.com',
    port: 443,
    apiv: 'v1.0',
    user: null,
    pass: null
};

Client.prototype = {

    get: function (path, cb) {
        var req = request({
            hostname: this.host,
            port    : this.port,
            method  : 'GET',
            path    : path && path.substring(0, 'https:'.length) === 'https:' ? path : this.path + path || '',
            auth    : this.user + ':' + this.pass
        }, function (res) {
            var data = '';
            res.setEncoding('utf8');
            res.on('data', function (chunk) {
                data += chunk;
            });
            res.on('end', function () {
                var res;
                try {
                    res = parseJSON( data );
                } catch (e) {
                    res = data;
                }
                cb( res );
            });
        });
        req.end();
        req.on('error', function (err) {
            throw new Error(err);
        });
    },

    post: function (path, data, cb) {
        var body = typeof data === 'object' ? serialize(data) : data;
        var req  = request({
            hostname: this.host,
            port    : this.port,
            method  : 'POST',
            path    : path && path.substring(0, 'https:'.length) === 'https:' ? path : this.path + path || '',
            auth    : this.user + ':' + this.pass
        }, function (res) {

            var data = '';

            res.setEncoding('utf8');
            res.on('data', function (chunk) {
                data += chunk;
            });
            res.on('end', function () {
                var res;
                try {
                    res = parseJSON( data );
                } catch (e) {
                    res = data;
                }
                cb( res );
            });

        });
        req.end(body);
        req.on('error', function (err) {
            throw new Error(err);
        });
    }
};

// ============================================================================
// Exports
// ============================================================================

module.exports = {
    Client: Client
};

