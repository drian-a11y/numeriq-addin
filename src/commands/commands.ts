/**
 * Commands - Handle ribbon button commands
 */

/* global Office */

Office.onReady(() => {
  // Commands are ready
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Numeriq command executed",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// Register the function
(Office as any).actions = {
  action
};
