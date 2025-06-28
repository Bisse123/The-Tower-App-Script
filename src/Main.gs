function doGet() {
  return HtmlService.createTemplateFromFile('pickerTest')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getClientId() {
  return '832137601831-t18ihai9kd0qmh2ko3mmn5t1p6iop0jn.apps.googleusercontent.com';
}

function getApiKey() {
  return 'AIzaSyBoleujCzXxUUhPoKmPzlbsOh0LEoer-gA';
}

function getAppId() {
  return '832137601831';
}
