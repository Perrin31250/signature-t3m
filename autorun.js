/* === DEBUG === */
console.log("autorun.js chargé");

Office.onReady(() => {
  console.log("Office prêt");
});

/**
 * Event Outlook : nouveau message
 */
function onNewMessageCompose(event) {
  console.log("OnNewMessageCompose déclenché");

  const signatureHtml = `
    <br/>
    <div style="font-family: Arial; font-size: 11pt;">
      <strong>Jean Dupont</strong><br/>
      Consultant IT<br/>
      <strong>T3M</strong><br/>
      📞 01 23 45 67 89<br/>
      ✉️ jean.dupont@t3m.fr
      <hr/>
    </div>
  `;

  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    function (result) {

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Évite les doublons
        if (!result.value.includes("Jean Dupont")) {
          Office.context.mailbox.item.body.setAsync(
            result.value + signatureHtml,
            { coercionType: Office.CoercionType.Html },
            function () {
              console.log("Signature insérée");
              event.completed();
            }
          );
        } else {
          console.log("Signature déjà présente");
          event.completed();
        }
      } else {
        console.error("Erreur lecture body");
        event.completed();
      }
    }
  );
}

/* === OBLIGATOIRE === */
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
