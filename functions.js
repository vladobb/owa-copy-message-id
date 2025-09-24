Office.onReady(() => {
  Office.actions.associate("copyMessageId", async () => {
    try {
      const item = Office.context.mailbox?.item;
      const id = item?.internetMessageId;
      if (!id) {
        return showInfobar("Open a message first.");
      }
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(id);
        showInfobar("Message-ID copied to clipboard.");
      } else {
        Office.context.ui.displayDialogAsync(
          `data:text/html,${encodeURIComponent(`
            <div style="font:14px/1.35 system-ui;padding:12px;max-width:480px">
              <div>Message-ID:</div>
              <pre style="white-space:pre-wrap;border:1px solid #ccc;padding:8px;border-radius:6px">${id}</pre>
              <button onclick="navigator.clipboard&&navigator.clipboard.writeText('${id.replace(/'/g,"&#39;")}');window.close()">Copy & Close</button>
            </div>
          `)}`,
          { height: 40, width: 50 }
        );
      }
    } catch (e) {
      showInfobar("Error: " + e.message);
    }
  });
});

function showInfobar(message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync("copyMID", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "icon16",
      persistent: false
    });
  } catch { /* ignore */ }
}
