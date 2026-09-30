/* global Office */
(function () {
  "use strict";

  var TARGET_URL = "https://chatgpt.com/g/g-p-6abc5b8c534c8191bef0b63fc87c478d/project";

  function launchJarvis(event) {
    try {
      Office.context.ui.openBrowserWindow(TARGET_URL);
    } finally {
      event.completed();
    }
  }

  Office.actions.associate("launchJarvis", launchJarvis);
}());
