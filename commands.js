/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  try {
    const targetUrl = "https://forms.office.com/e/0WMwRUR02J";

    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType
        .InformationalMessage,
      message: "Opened Microsoft Form in browser.",
      icon: "Icon.80x80", // Must match the icon id in your manifest resources
      persistent: true,
    };

    // Open Microsoft Form in a browser window
    window.open(targetUrl);

    // Optionally show a notification in the current item, if available
    if (
      Office.context &&
      Office.context.mailbox &&
      Office.context.mailbox.item
    ) {
      Office.context.mailbox.item.notificationMessages.replaceAsync(
        "formAction",
        message
      );
    }
  } finally {
    // Be sure to indicate when the add-in command function is complete
    event.completed();
  }
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
