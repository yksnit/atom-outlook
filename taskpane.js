(() => {
  const templateHtml = [
    "<p>Hi there,</p>",
    "<p>Thanks for reaching out. I’ll get back to you shortly with details.</p>",
    "<p>Best,<br>[Your name]</p>"
  ].join("");

  const status = document.getElementById("status");

  function setStatus(message) {
    status.textContent = message || "";
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      setStatus("This add-in runs in Outlook.");
      return;
    }

    const button = document.getElementById("insert-template");
    if (button) {
      button.addEventListener("click", insertTemplate);
    }

    setStatus("Ready in compose mode.");
  });

  function insertTemplate() {
    const item = Office.context.mailbox.item;
    if (!item || !item.body || typeof item.body.setSelectedDataAsync !== "function") {
      setStatus("Open a draft message to use the template.");
      return;
    }

    item.body.setSelectedDataAsync(
      templateHtml,
      { coercionType: Office.CoercionType.Html },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Template inserted into your draft.");
        } else {
          setStatus(asyncResult.error?.message || "Could not insert template.");
        }
      }
    );
  }
})();
