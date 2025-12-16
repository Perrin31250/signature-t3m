/* global Office, OfficeRuntime */

// Association de la fonction pour Outlook Windows Classique
if (typeof Office !== 'undefined' && Office.actions) {
    Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
}

const TEMPLATE = `
<table cellpadding="0" cellspacing="0" style="font-family: Arial, sans-serif; font-size: 12px; color:#000000; line-height:1.4; border-top: 1px solid #eeeeee; margin-top: 20px; padding-top: 20px;">
  <tr>
    <td style="padding-right:16px; border-right:1px solid #000000; vertical-align:top;">
      <a href="https://www.groupet3m.com">
        <img src="https://recrutement.groupet3m.com/wp-content/uploads/2025/11/T3M_N_LAVAIL_C-1.png" width="120" alt="T3M" style="display:block; border:0;">
      </a>
    </td>
    <td style="padding-left:16px; vertical-align:top;">
      <div style="font-size:14px; font-weight:bold; color:#000000; margin-bottom:2px;">{{DisplayName}}</div>
      <div style="margin-bottom:2px; color:#666666;">{{JobTitle}}</div>
      <div style="font-weight:bold; margin-bottom:4px;">{{CompanyLegal}}</div>
      {{#Mobile}}<div style="margin-bottom:2px;">Tél : {{Mobile}}</div>{{/Mobile}}
      <div style="margin-top:8px;"><a href="https://www.groupet3m.com" style="color:#000000; text-decoration:none; font-weight:bold;">www.groupet3m.com</a></div>
    </td>
  </tr>
</table>`;

async function onNewMessageCompose(event) {
    try {
        // 1. Désactiver la signature par défaut de l'utilisateur
        await new Promise((resolve) => {
            Office.context.mailbox.item.disableClientSignatureAsync(resolve);
        });

        // 2. Récupérer les données de base (Profil local Outlook)
        const userProfile = Office.context.mailbox.userProfile;
        
        // 3. Essayer de récupérer des infos plus complètes via Graph (optionnel)
        let graphData = null;
        try {
            const token = await getGraphToken();
            if (token) {
                graphData = await getGraphMe(token);
            }
        } catch (e) {
            console.log("Graph non disponible, utilisation profil local.");
        }

        // 4. Fusion des données
        const data = {
            DisplayName: graphData?.displayName || userProfile.displayName || "Collaborateur T3M",
            JobTitle: graphData?.jobTitle || "",
            Mobile: graphData?.mobilePhone || "",
            CompanyLegal: graphData?.companyName || "GROUPE T3M"
        };

        // 5. Générer le HTML et insérer la signature
        const htmlSignature = applyTemplate(TEMPLATE, data);
        
        Office.context.mailbox.item.body.setSignatureAsync(
            htmlSignature,
            { coercionType: Office.CoercionType.Html },
            () => {
                event.completed(); // Signale qu'on a fini
            }
        );

    } catch (error) {
        console.error("Erreur signature:", error);
        event.completed();
    }
}

async function getGraphToken() {
    try {
        // Tente de récupérer un token SSO (nécessite App Registration dans Azure)
        return await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
    } catch (e) {
        return null;
    }
}

async function getGraphMe(token) {
    if (!token) return null;
    const res = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mobilePhone,companyName", {
        headers: { "Authorization": `Bearer ${token}` }
    });
    return res.ok ? res.json() : null;
}

function applyTemplate(tpl, data) {
    let out = tpl;
    // Gérer les blocs conditionnels {{#Key}}...{{/Key}}
    out = out.replace(/{{#(\w+)}}([\s\S]*?){{\/\1}}/g, (m, key, inner) => data[key] ? inner : "");
    // Remplacer les variables {{Key}}
    out = out.replace(/{{(\w+)}}/g, (m, key) => data[key] ?? "");
    return out;
}
