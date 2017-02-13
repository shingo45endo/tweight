'use strict';

function TwitterWebService(key, secret, authCallback) {
  if (!key || !secret || !authCallback) {
    throw new TypeError('Invalid arguments');
  }
  if (!(this instanceof TwitterWebService)) {
    throw new TypeError('TwitterWebService is a constructor');
  }

  var consumerKey = key;
  var consumerSecret = secret;
  var authCallbackName = (typeof authCallback === 'function') ? authCallback.name : authCallback;

  /**
   * Configures the service.
   */
  this.getService = function() {
    return OAuth1.createService('Twitter')
    .setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
    .setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
    .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')
    .setConsumerKey(consumerKey)
    .setConsumerSecret(consumerSecret)
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

  /**
   * Tweets a text.
   */
  this.tweet = function(text, mediaId) {
    // Creates the OAuth1 service for Twitter API.
    var service = this.getService();
    if (!service.hasAccess()) {
      Logger.log('Have not been authorized.');
      return;
    }

    // Makes the options for the request.
    var options = {
      method: 'post',
      escaping: false
    };
    var encodedText = encodeURIComponent(text).replace(/[!'()*]/g, function(ch) {return "%" + ch.charCodeAt(0).toString(16);});
    if (mediaId) {
      encodedText += '&media_ids=' + mediaId;
    }
    options.payload = 'status=' + encodedText;

    // Accesses to the Twitter API to tweet.
    var response = service.fetch('https://api.twitter.com/1.1/statuses/update.json', options);
    if (response.getResponseCode() >= 400) {
      Logger.log('Something wrong with HTTP response:');
      Logger.log(response);
      return;
    }

    // Checks the response from Twitter API.
    var result = JSON.parse(response.getContentText());
    if (false) {
      Logger.log('Something wrong with Twitter API: %s', response.getContentText());
      return;
    }
  };

  /**
   * Tweets a text with media.
   */
  this.tweetWithMedia = function(text, blob) {
    // Creates the OAuth1 service for Twitter API.
    var service = this.getService();
    if (!service.hasAccess()) {
      Logger.log('Have not been authorized.');
      return;
    }

    // Makes a payload including the media data.
    var boundary = '' + (new Date()).getTime();
    var requestBody = Utilities.newBlob(
      '--' + boundary + '\r\n' +
      'Content-Disposition: form-data; name="media"; filename=""\r\n' +
      'Content-Type: ' + blob.getContentType() + '\r\n\r\n').getBytes();
    requestBody = requestBody.concat(blob.getBytes());
    requestBody = requestBody.concat(Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes());

    // Makes the options for the request.
    var options = {
      method: 'post',
      contentType: 'multipart/form-data; boundary=' + boundary,
      payload: requestBody,
    };

    // Accesses to the Twitter API to upload.
    var response = service.fetch('https://upload.twitter.com/1.1/media/upload.json', options);
    if (response.getResponseCode() >= 400) {
      Logger.log('Something wrong with HTTP response:');
      Logger.log(response);
      return;
    }

    // Checks the response from Twitter API.
    var result = JSON.parse(response.getContentText());
    var media_id_string = result.media_id_string;
    if (!media_id_string) {
      Logger.log('Something wrong with Twitter API: %s', response.getContentText());
      return;
    }

    // Tweets the text with the uploaded media.
    this.tweet(text, media_id_string);
  };
}
