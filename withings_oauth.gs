'use strict';

var SERVICE_NAME = 'Withings';

/**
 * Authorizes and makes a request to the Withings API.
 */
function run() {
  var service = getWithingsService();
  if (!service.hasAccess()) {
    showSidebar();
    return;
  }

  var userid = PropertiesService.getScriptProperties().getProperty('userid');
  var url = 'https://wbsapi.withings.net/measure?action=getmeas&userid=' + userid;
  var response = service.fetch(url);
  var result = JSON.parse(response.getContentText());

  Logger.log(result);
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getWithingsService();
  service.reset();
}

/**
 * Configures the service.
 */
function getWithingsService() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return OAuth1.createService(SERVICE_NAME)
    .setAccessTokenUrl('https://oauth.withings.com/account/access_token')
    .setRequestTokenUrl('https://oauth.withings.com/account/request_token')
    .setAuthorizationUrl('https://oauth.withings.com/account/authorize')
    .setConsumerKey(scriptProperties.getProperty('withingConsumerKey'))
    .setConsumerSecret(scriptProperties.getProperty('withingConsumerSecret'))
    .setParamLocation('uri-query')
    .setOAuthVersion('1.0')
    .setCallbackFunction('authCallback')
    .setPropertyStore(scriptProperties);
}

/**
 * Logs the callback URL to register.
 */
function logCallbackUrl() {
  var service = getWithingsService();
  Logger.log(service.getCallbackUrl());
}

/**
 * Shows the sidebar to display the authorization URL.
 */
function showSidebar() {
  var service = getWithingsService();
  if (!service.hasAccess()) {
    var authorizationUrl = service.authorize();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
        'Reopen the sidebar when the authorization is complete.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput('Already authorized.'));
  }
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getWithingsService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
