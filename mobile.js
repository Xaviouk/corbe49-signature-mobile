/*
 * Test prudent Outlook Mobile - Android.
 * Ne désactive pas la signature native dans les paramètres Outlook.
 * À la création d'un message, remplace uniquement la signature visible dans ce brouillon
 * par une signature de test. Ne pas envoyer de message de test à un client.
 */
Office.onReady();

function onNewMessageComposeHandler(event) {
  const item = Office.context.mailbox.item;
  const testSignature =
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;padding-top:8px;border-top:2px solid #c8102e;">' +
    '<strong style="color:#c8102e;">CORBE — TEST ANDROID</strong><br>' +
    '<span>Complément de signature en cours de validation</span><br>' +
    '<span style="font-size:11px;color:#777;">Ne pas envoyer — message de test</span>' +
    '</div>';

  item.body.setSignatureAsync(
    testSignature,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Échec ajout signature test : " + result.error.message);
      }
      event.completed();
    }
  );
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
