/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function toggleProtection(args) {
  try {
    // eslint-disable-next-line no-undef
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      sheet.load("protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect();
      } else {
        sheet.protection.protect();
      }

      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    Console.error(error);
  }

  args.completed();
}
Office.actions.associate("toggleProtection", toggleProtection);

let dialog = null;
async function functionsToolset(args) {
  try {
    Office.context.ui.displayDialogAsync(
      // eslint-disable-next-line no-undef
      document.location.origin + "/" + "toolset.html",
      { height: 45, width: 25 },
      function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    );
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    Console.error(error);
  }
  args.completed();
}
function processMessage(arg) {
  // eslint-disable-next-line no-undef
  document.getElementById("user-name1").innerHTML = arg.message;
  dialog.close();
}
Office.actions.associate("functionsToolset", functionsToolset);

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

// the add-in command functions need to be available in global scope
g.action = action;
