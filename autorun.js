/* global Office */

// ⚠️ Sur Windows classique, Office.onReady/initialize ne s'exécutent pas dans le runtime JS.
// --> Associe au niveau global + garde au besoin onReady pour Web/Mac/Nouveau Outlook.
if (Office && Office.actions) {
  Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
}

// Optionnel : utile pour Web/Mac/Nouveau Outlook
Office.onReady(() => {
  // Rien d'obligatoire ici pour les events, l'association ci-dessus suffit.
});

const TEMPLATE = `<!-- Signature avec placeholders -->
<table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 12px; color:#000000; line-height:1.4;">
  <tr>
    <td style="padding-right:16px; border-right:1px solid #000000;">
      https://www.groupet3m.com
        https://recrutement.groupet3m.com/wp-content/uploads/2025/11/T3M_N_LAVAIL_C-1.png
      </a>
    </td>
    <td style="padding-left:16px;">
      <div style="font-size:14px; font-weight:bold; margin-bottom:2px;">{{DisplayName}}</div>
      <div style="margin-bottom:2px;">{{JobTitle}}</div>
      <div style="font-weight:bold; margin-bottom:4px;">{{CompanyLegal}}</div>
      {{#Mobile}}<div style="margin-bottom:2px;">{{Mobile}}</div>{{/Mobile}}
      <div style="margin-bottom:8px;">https://www.groupet3m.comwww.groupet3m.com</a></div>
    </td>
  </tr>
</table>`;

async function onNewMessageCompose(event) {
  try {
    // Désactive la signature client Outlook si activée
    Office.context.mailbox.item.isClientSignatureEnabledAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
        Office.context.mailbox.item.disableClientSignatureAsync(() => {});
      }
    });

    const accessToken = await getGraphToken();  // TODO: implémenter SSO (MSAL/OfficeRuntime.auth)
    const profile = await getGraphMe(accessToken);
    const p = Office.context.mailbox.userProfile;

    const data = {
      DisplayName: profile?.displayName ?? p.displayName ?? "",
      JobTitle: profile?.jobTitle ?? "",
      Mobile: profile?.mobilePhone ?? "",
      CompanyLegal: profile?.companyName ?? "GROUPE T3M"
    };

    const html = applyTemplate(TEMPLATE, data);
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      () => event.completed()
    );
  } catch (e) {
    // Toujours terminer l’événement
    event.completed();
  }
}

async function getGraphMe(token) {
  if (!token) return null;
  const res = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mobilePhone,companyName", {
    headers: { Authorization: `Bearer ${token}` }
  });
  return res.ok ? res.json() : null;
}

async function getGraphToken() {
  // À implémenter : OfficeRuntime.auth ou MSAL (selon ta stratégie)
  return null;
}

function applyTemplate(tpl, data) {
  let out = tpl  let out = tpl;
  // Sections conditionnelles {{#Key}}...{{/Key}}
  out = out.replace(/{{#(\w+)}}([\s\S]*?){{\/\1}}/g, (m, key, inner) => data[key] ? inner : "");
  // Placeholders {{Key}}
  out = out.replace(/{{(\w+)}}/g, (m, key) => data[key] ?? "");
  return out;

}
