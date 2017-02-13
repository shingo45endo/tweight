'use strict';

var WithingsWebService = function(key, secret, authCallback) {
  if (!key || !secret || !authCallback) {
    throw new TypeError('Invalid arguments');
  }
  if (!(this instanceof WithingsWebService)) {
    throw new TypeError('WithingsWebService is a constructor');
  }

  var consumerKey = key;
  var consumerSecret = secret;
  var authCallbackName = (typeof authCallback === 'function') ? authCallback.name : authCallback;

  /**
   * Configures the service.
   */
  this.getService = function() {
    return OAuth1.createService('Withings')
    .setAccessTokenUrl('https://oauth.withings.com/account/access_token')
    .setRequestTokenUrl('https://oauth.withings.com/account/request_token')
    .setAuthorizationUrl('https://oauth.withings.com/account/authorize')
    .setConsumerKey(consumerKey)
    .setConsumerSecret(consumerSecret)
    .setParamLocation('uri-query')
    .setOAuthVersion('1.0')
    .setCallbackFunction(authCallbackName)
    .setPropertyStore(PropertiesService.getScriptProperties());
  };

  /**
   * Handles the OAuth callback.
   */
  this.authCallback = function(request) {
    var service = this.getService();
    var isAuthorized = service.handleCallback(request);
    if (isAuthorized) {
      return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
      return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
  };

  /**
   * Reset the authorization state, so that it can be re-tested.
   */
  this.reset = function() {
    var service = this.getService();
    service.reset();
  };

  /**
   * Logs the callback URL to register.
   */
  this.logCallbackUrl = function() {
    var service = this.getService();
    Logger.log(service.getCallbackUrl());
  };

  /**
   * Logs the authorization URL.
   */
  this.logAuthorizationUrl = function() {
    var service = this.getService();
    if (!service.hasAccess()) {
      Logger.log('Open the following URL and re-run the script: %s', service.authorize());
    } else {
      Logger.log('Already authorized.');
    }
  };
};
