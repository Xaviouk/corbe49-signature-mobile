/* CORBÉ 49 — Signature mobile dynamique
 * Récupère le profil Microsoft 365 via Microsoft Graph avec MSAL Nested App Authentication.
 * En cas d'échec d'authentification, conserve les informations Outlook disponibles sans inventer de fonction ni de téléphone.
 */
Office.onReady();

const CORBE_AUTH = {
  clientId: "11d4d023-0187-483e-aa31-9b3f55153aee",
  authority: "https://login.microsoftonline.com/75b354dc-0810-4caa-90a3-ac64ef88c77a",
  redirectUri: "https://xaviouk.github.io/corbe49-signature-mobile/auth.html"
};
let corbePcaPromise;
async function ensureMsalLoaded() {
  if (window.msal && window.msal.createNestablePublicClientApplication) return;
  // Le runtime d’événement Outlook peut exécuter mobile.js sans avoir exécuté
  // les balises <script> de mobile.html. Charger MSAL explicitement dans ce runtime.
  if (typeof document === "undefined") {
    throw new Error("Impossible de charger MSAL : aucun document disponible dans le runtime Outlook.");
  }
  await new Promise((resolve, reject) => {
    const existing = document.querySelector('script[data-corbe-msal="true"]');
    if (existing) {
      if (window.msal && window.msal.createNestablePublicClientApplication) return resolve();
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", () => reject(new Error("Échec du chargement de MSAL Browser.")), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = "https://alcdn.msauth.net/browser/4.26.0/js/msal-browser.min.js";
    script.async = true;
    script.dataset.corbeMsal = "true";
    script.onload = resolve;
    script.onerror = () => reject(new Error("Échec du chargement de MSAL Browser."));
    document.head.appendChild(script);
  });
  if (!window.msal || !window.msal.createNestablePublicClientApplication) {
    throw new Error("MSAL a été chargé, mais NAA n’est pas disponible dans cette version.");
  }
}
async function getCorbePca() {
  await ensureMsalLoaded();
  if (!corbePcaPromise) {
    corbePcaPromise = window.msal.createNestablePublicClientApplication({
      auth: {
        clientId: CORBE_AUTH.clientId,
        authority: CORBE_AUTH.authority,
        redirectUri: CORBE_AUTH.redirectUri
      },
      cache: { cacheLocation: "localStorage" }
    });
  }
  return corbePcaPromise;
}

const corbeSignatureTemplate = String.raw`<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>Signature2026Corbe49</title>
</head>
<body style="margin:0;padding:0;background:transparent;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">

<!-- V53 = V52 + largeur de la photo verrouillée (cellule et image : 78px minimum et maximum) : sur mobile, la photo était réduite et écrasée car c'est la seule colonne qui pouvait rétrécir.
     V52 = V51 + espace insécable entre prénom et nom (le nom ne passe plus sur deux lignes sur mobile) + blocage de la mise à l'échelle du texte.
     V51 = V50 + hauteurs de ligne exactes sur les images et les textes des coordonnées (Outlook ajoutait de la hauteur sous les images et autour des textes).
     V50 : CONSTRUCTION SANS CELLULE FINE
     Outlook classique donne à toute cellule de tableau la hauteur d'un paragraphe normal (environ 19px) :
     les espaceurs de 2 à 11px et les traits dessinés avec une cellule colorée devenaient de gros blocs.
     Ici :
       - plus aucun espaceur : les espaces sont des marges internes (padding), qu'Outlook respecte à l'identique ;
       - plus aucun trait en cellule colorée : les traits sont des BORDURES (ils restent fins même si la cellule est plus haute) ;
       - traits verticaux : bordure droite du bloc qu'ils accompagnent (identité pour le rouge, coordonnées pour le gris),
         donc ils suivent la hauteur de leur bloc ; la marge interne de 8px au-dessus et en dessous les laisse en retrait.
     Repères : haut du nom, de l'e-mail, de l'adresse et des traits = 8px ; bas des traits, du logo CORBÉ 49, des pictos et du GIF = 100px. -->
<table cellpadding="0" cellspacing="0" border="0"
       style="table-layout:fixed;mso-table-lspace:0pt;mso-table-rspace:0pt;
               -webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;
               border-collapse:collapse;background:transparent;
               font-family:'Trebuchet MS',Arial,Helvetica,sans-serif;color:#202c39;">
<tr>

<!-- PHOTO OVALE V57 : dimensions strictement verrouillées pour Outlook mobile -->
<td width="78" valign="top" nowrap
    style="width:78px;min-width:78px;max-width:78px;padding:0 10px 0 20px;white-space:nowrap;">
  <div style="font-size:0;line-height:0;width:78px;height:108px;overflow:hidden;">
    <img src="{{PHOTO_URL}}"
         alt="{{DISPLAY_NAME_HTML}}" width="78" height="108"
         style="display:block;width:78px;height:108px;border:0;margin:0;padding:0;
                max-width:none;max-height:none;min-width:0;min-height:0;
                -ms-interpolation-mode:bicubic;">
  </div>
</td>

<!-- IDENTITÉ + TRAIT ROUGE (bordure droite) : 8px de retrait en haut et en bas -->
<td valign="top" style="padding:8px 0 8px 0;">
  <table cellpadding="0" cellspacing="0" border="0"
         style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
    <tr>
      <td valign="top" style="padding:0 10px 0 0;border-right:2px solid #E30613;">

        <table cellpadding="0" cellspacing="0" border="0"
               style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">

          <tr>
            <td valign="top" style="padding:0;">
              <div style="font-family:'Trebuchet MS',Arial,Helvetica,sans-serif;
                          font-size:18px;line-height:18px;mso-line-height-rule:exactly;font-weight:800;
                          white-space:nowrap;color:#202c39;margin:0;">
                {{DISPLAY_NAME_HTML}}
              </div>
            </td>
          </tr>

          <tr>
            <td valign="top" style="padding:6px 0 0 0;">
              <div style="font-family:'Trebuchet MS',Arial,Helvetica,sans-serif;
                          font-size:10px;line-height:12px;mso-line-height-rule:exactly;font-weight:400;
                          letter-spacing:2px;color:#3e4650;margin:0 0 0 2px;white-space:nowrap;">
                {{JOB_TITLE_HTML}}
              </div>
            </td>
          </tr>

          <!-- Soulignement rouge : bordure haute d'une cellule de 39px (6px d'espace au-dessus, environ 11px en dessous) -->
          <tr>
            <td valign="top" style="padding:6px 0 0 0;">
              <table cellpadding="0" cellspacing="0" border="0" width="39"
                     style="width:39px;mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
                <tr>
                  <td width="39" height="11" valign="top"
                      style="width:39px;height:11px;padding:0;border-top:2px solid #E30613;font-size:1px;line-height:1px;mso-line-height-rule:exactly;">
                    <div style="font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td valign="bottom" style="padding:0;">
              <a href="https://corbe-cuisine-pro.com" target="_blank" style="text-decoration:none;">
                <div style="font-size:1px;line-height:37px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/logo-corbe-agence-49.png"
                     alt="CORBÉ 49" width="141" height="37"
                     style="display:block;width:141px;height:37px;border:0;margin:0;"></div>
              </a>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</td>

<!-- COORDONNÉES + TRAIT GRIS (bordure droite) : 10px après le trait rouge, 8px de retrait en haut et en bas -->
<td valign="top" style="padding:8px 0 8px 10px;">
  <table cellpadding="0" cellspacing="0" border="0"
         style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
    <tr>
      <td valign="top" height="92" style="height:92px;padding:0 10px 0 0;border-right:2px solid #c9c9c9;">

        <table cellpadding="0" cellspacing="0" border="0"
               style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">

          <tr>
            <td width="18" valign="middle" style="width:18px;padding:0 0 11px 0;">
              <div style="font-size:1px;line-height:14px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-enveloppe-signa.png"
                   alt="Email" width="14" style="display:block;width:14px;height:auto;border:0;"></div>
            </td>
            <td valign="middle" style="padding:0 0 12px 0;font-size:10px;line-height:1.1;
                                        font-weight:400;white-space:nowrap;">
              <div style="font-size:10px;line-height:12px;mso-line-height-rule:exactly;white-space:nowrap;"><a href="mailto:{{EMAIL_ATTR}}" style="color:#202c39;text-decoration:none;">
                {{EMAIL_HTML}}
              </a></div>
            </td>
          </tr>

          <tr>
            <td valign="middle" style="padding:0 0 11px 0;">
              <div style="font-size:1px;line-height:14px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-globe-signa.png"
                   alt="Web" width="14" style="display:block;width:14px;height:auto;border:0;"></div>
            </td>
            <td valign="middle" style="padding:0 0 12px 0;font-size:10px;line-height:1.1;
                                        font-weight:400;white-space:nowrap;">
              <div style="font-size:10px;line-height:12px;mso-line-height-rule:exactly;white-space:nowrap;"><a href="https://corbe-cuisine-pro.com" target="_blank"
                 style="color:#202c39;text-decoration:none;">
                corbe-cuisine-pro.com
              </a></div>
            </td>
          </tr>

          <tr>
            <td valign="middle" style="padding:0 0 11px 0;">
              <div style="font-size:1px;line-height:14px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-telephoni-signa.png"
                   alt="Téléphone" width="14" style="display:block;width:14px;height:auto;border:0;"></div>
            </td>
            <td valign="middle" style="padding:0 0 11px 0;font-size:10px;line-height:1.1;
                                        font-weight:700;white-space:nowrap;">
              <div style="font-size:10px;line-height:12px;mso-line-height-rule:exactly;white-space:nowrap;"><a href="tel:{{PHONE_DIGITS}}" style="color:#E30613;text-decoration:none;">
                {{PHONE_HTML}}
              </a></div>
            </td>
          </tr>

          <tr>
            <td valign="middle" style="padding:0;">
              <div style="font-size:1px;line-height:14px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-fichecontact-signa.png"
                   alt="Contact" width="14" style="display:block;width:14px;height:auto;border:0;"></div>
            </td>
            <td valign="middle" style="padding:0;font-size:10px;line-height:1.1;font-weight:700;white-space:nowrap;">
              <div style="font-size:10px;line-height:12px;mso-line-height-rule:exactly;white-space:nowrap;"><a href="https://www.corbe-cuisine-pro.com/vcf/49/{{EMAIL_ATTR}}.vcf" target="_blank" style="color:#E30613;text-decoration:none;">
                Ajouter ce contact
              </a></div>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</td>

<!-- BLOC DROIT : 10px après le trait gris, marge droite 20px ; 8px de retrait en haut -->
<td valign="top" style="padding:8px 20px 0 10px;">
  <table cellpadding="0" cellspacing="0" border="0"
         style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">

    <!-- ADRESSE (3 lignes x 13px) : 1re ligne alignée sur l'e-mail -->
    <tr>
      <td valign="top" style="padding:0;">
        <table cellpadding="0" cellspacing="0" border="0" width="144"
               style="width:144px;mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
          <tr>
            <td width="18" valign="top" style="width:18px;padding:0;">
              <div style="font-size:1px;line-height:14px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-gps-signa.png"
                   alt="Adresse" width="14"
                   style="display:block;width:14px;height:auto;border:0;"></div>
            </td>
            <td width="126" valign="top"
                style="width:126px;padding:0;font-family:'Trebuchet MS',Arial,Helvetica,sans-serif;
                       font-size:10px;line-height:13px;mso-line-height-rule:exactly;font-weight:400;
                       color:#202c39;white-space:nowrap;">
              <div style="font-size:10px;line-height:13px;mso-line-height-rule:exactly;"><a href="https://www.google.com/maps/search/?api=1&query=22%20rue%20de%20l%27Europe%2C%20ZI%20Les%20Trois%20Routes%2C%2049120%20Chemill%C3%A9-en-Anjou"
                 target="_blank" rel="noopener"
                 style="color:#202c39;text-decoration:none;">
                22 rue de l’Europe<br>
                ZI Les Trois Routes<br>
                49120 CHEMILLÉ-EN-ANJOU
              </a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- LIGNE GRISE + "Adhérent du" : 9px d'espace au-dessus ; ligne grise = bordure basse (80px), texte centré au-dessus du logo GIF -->
    <tr>
      <td valign="top" style="padding:9px 0 0 0;">
        <table cellpadding="0" cellspacing="0" border="0" width="144"
               style="width:144px;mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
          <tr>
            <td width="80" height="8" valign="bottom"
                style="width:80px;height:8px;padding:0;border-bottom:2px solid #c5c5c5;font-size:1px;line-height:1px;mso-line-height-rule:exactly;">
              <div style="font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</div>
            </td>
            <td width="27" valign="bottom" style="width:27px;padding:0;font-size:1px;line-height:1px;mso-line-height-rule:exactly;">
              <div style="font-size:1px;line-height:1px;mso-line-height-rule:exactly;">&nbsp;</div>
            </td>
            <td width="37" valign="bottom" align="center"
                style="width:37px;padding:0;white-space:nowrap;">
              <div style="font-family:'Trebuchet MS',Arial,Helvetica,sans-serif;
                          font-size:6px;line-height:8px;mso-line-height-rule:exactly;font-weight:400;
                          color:#706f6f;text-align:center;">
                Adhérent du
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- RÉSEAUX SOCIAUX + GIF : pictos et GIF calés en bas ; trait vertical = bordure gauche de la cellule du GIF -->
    <tr>
      <td valign="top" style="padding:0;">
        <table cellpadding="0" cellspacing="0" border="0" width="144"
               style="width:144px;mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
          <tr>
            <td width="97" valign="bottom" style="width:97px;padding:0;">
              <table cellpadding="0" cellspacing="0" border="0" width="90"
                     style="width:90px;mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;">
                <tr>
                  <td width="33" style="width:33px;padding:0;">
                    <a href="https://www.facebook.com/corbecuisineprofessionnelle/" target="_blank">
                      <div style="font-size:1px;line-height:23px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-fac-signa.png"
                           alt="Facebook" width="22"
                           style="display:block;width:22px;height:auto;border:0;"></div>
                    </a>
                  </td>
                  <td width="33" style="width:33px;padding:0;">
                    <a href="https://fr.linkedin.com/company/corb%C3%A9-cuisine-professionnelle" target="_blank">
                      <div style="font-size:1px;line-height:23px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-link-signa.png"
                           alt="LinkedIn" width="22"
                           style="display:block;width:22px;height:auto;border:0;"></div>
                    </a>
                  </td>
                  <td width="24" style="width:24px;padding:0;">
                    <a href="https://www.youtube.com/@corbecuisineprofessionnell9733" target="_blank">
                      <div style="font-size:1px;line-height:23px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/picto-you-signa.png"
                           alt="YouTube" width="22"
                           style="display:block;width:22px;height:auto;border:0;"></div>
                    </a>
                  </td>
                </tr>
              </table>
            </td>

            <td width="37" valign="bottom" align="center"
                style="width:37px;padding:0 0 0 9px;border-left:1px solid #c9c9c9;">
              <a href="https://www.groupegif.com/" target="_blank" style="text-decoration:none;">
                <div style="font-size:1px;line-height:36px;mso-line-height-rule:exactly;"><img src="https://www.corbe-cuisine-pro.com/wp-content/uploads/logo-gif.png"
                     alt="GIF – Groupement des Installateurs Français" width="35"
                     style="display:block;width:35px;height:auto;border:0;"></div>
              </a>
            </td>
          </tr>
        </table>
      </td>
    </tr>

  </table>
</td>

</tr>
</table>

</body>
</html>
`;

function escapeHtml(value) {
  return String(value || "").replace(/[&<>\"]/g, ch => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[ch]));
}
function digits(value) { return String(value || "").replace(/\D/g, ""); }
function formatPhone(value) {
  const d = digits(value);
  if (d.length === 10) return d.match(/.{1,2}/g).join(" ");
  if (d.length === 11 && d.startsWith("33")) return "0" + d.slice(2).match(/.{1,2}/g).join(" ");
  return value || "";
}
function splitName(displayName, email) {
  const cleaned = String(displayName || "").trim();
  if (cleaned) return cleaned;
  const local = String(email || "").split("@")[0].replace(/[._-]+/g, " ").trim();
  return local.replace(/\b\w/g, c => c.toUpperCase());
}
async function getGraphProfile() {
  try {
    const pca = await getCorbePca();
    const result = await pca.acquireTokenSilent({ scopes: ["User.Read"] });
    const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail,userPrincipalName,mobilePhone,businessPhones", {
      headers: { Authorization: "Bearer " + result.accessToken }
    });
    if (!response.ok) throw new Error("Microsoft Graph a répondu " + response.status);
    return await response.json();
  } catch (e) {
    // Un événement de composition ne doit pas ouvrir de fenêtre interactive. La signature
    // reste utilisable avec les champs Outlook disponibles, et les champs Graph restent vides.
    console.warn("Profil Graph indisponible en mode silencieux.", e);
    return null;
  }
}
function renderSignature(profile, graph) {
  const email = String((graph && (graph.mail || graph.userPrincipalName)) || profile.emailAddress || "").trim().toLowerCase();
  const displayName = splitName((graph && graph.displayName) || profile.displayName, email);
  const jobTitle = String((graph && graph.jobTitle) || "").trim();
  const phoneRaw = String((graph && graph.mobilePhone) || (graph && graph.businessPhones && graph.businessPhones[0]) || "").trim();
  const phone = formatPhone(phoneRaw);
  const phoneDigits = digits(phoneRaw);
  const photoUrl = email ? "https://www.corbe-cuisine-pro.com/wp-content/uploads/" + encodeURIComponent(email) + "_.png" : "https://www.corbe-cuisine-pro.com/wp-content/uploads/xavier.gaillard@corbe49.fr_.png";
  const replacements = {
    "{{PHOTO_URL}}": photoUrl,
    "{{DISPLAY_NAME_HTML}}": escapeHtml(displayName).replace(/ /g, "&nbsp;"),
    "{{JOB_TITLE_HTML}}": escapeHtml(jobTitle),
    "{{EMAIL_ATTR}}": email,
    "{{EMAIL_HTML}}": escapeHtml(email),
    "{{PHONE_DIGITS}}": phoneDigits,
    "{{PHONE_HTML}}": escapeHtml(phone)
  };
  let html = corbeSignatureTemplate;
  for (const [key, value] of Object.entries(replacements)) html = html.split(key).join(value);
  return html;
}

function onNewMessageComposeHandler(event) {
  const item = Office.context.mailbox.item;
  const profile = Office.context.mailbox.userProfile || {};
  getGraphProfile().then(graph => {
    const signatureHtml = renderSignature(profile, graph);
    item.body.setSignatureAsync(signatureHtml, { coercionType: Office.CoercionType.Html }, result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Échec ajout signature CORBE 49 : " + result.error.message);
      }
      event.completed();
    });
  }).catch(error => {
    console.log("Erreur de préparation signature CORBE 49 :", error);
    const signatureHtml = renderSignature(profile, null);
    item.body.setSignatureAsync(signatureHtml, { coercionType: Office.CoercionType.Html }, result => {
      if (result.status === Office.AsyncResultStatus.Failed) console.log(result.error.message);
      event.completed();
    });
  });
}
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
