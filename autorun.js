// Cette fonction est appelée automatiquement par Outlook à chaque nouveau message
function onNewMessageCompose(event) {
    const signature = "<br><br>---<br><b>GROUPE T3M</b><br>Signature Automatique";
    
    // Insère la signature dans le corps du mail
    Office.context.mailbox.item.body.setSelectedDataAsync(
        signature,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Signature insérée avec succès.");
            } else {
                console.error("Erreur lors de l'insertion : " + asyncResult.error.message);
            }
            // Indique à Outlook que l'opération est terminée
            event.completed();
        }
    );
}

// Enregistrement de la fonction pour qu'Outlook la reconnaisse
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
