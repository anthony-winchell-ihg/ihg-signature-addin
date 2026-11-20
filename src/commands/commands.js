import { createNestablePublicClientApplication } from "@azure/msal-browser";
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

const msalConfig = {
  auth: {
    clientId: "0d99f9c6-d672-49ad-a78b-8a69b7d47798",
    authority: "https://login.microsoftonline.com/a9738f89-7274-4ea6-9905-2a2ea54d74cc",
  },
};

let pcaPromise;

async function getGraphToken() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication(msalConfig);
  }
  const pca = await pcaPromise;

  const result = await pca.acquireTokenSilent({ scopes: ["User.Read"] });
  return result.accessToken;
}

async function getMyProfile(token) {
  const url =
    "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,businessPhones,mobilePhone";

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    throw new Error(`Graph request failed: ${response.status}`);
  }
  return response.json();
}

async function onNewMessageComposeHandler(event) {
  try {
    const token = await getGraphToken();
    const profile = await getMyProfile(token);
    const signature = buildSignature(profile);

    Office.context.mailbox.item.disableClientSignatureAsync((disableResult) => {
      if (disableResult.status === Office.AsyncResultStatus.Failed) {
        console.error(disableResult.error);
      }

      Office.context.mailbox.item.body.setSignatureAsync(
        signature,
        { coercionType: Office.CoercionType.Html },
        (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Failed) {
            console.error(setResult.error);
          }

          event.completed();
        }
      );
    });
  } catch (err) {
    console.error(err);
    event.completed();
  }
}

function escapeHTML(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildSignature(profile) {
  const displayName = escapeHTML(profile.displayName);
  const jobTitle = escapeHTML(profile.jobTitle);
  const businessPhone = escapeHTML(profile.businessPhones[0]);
  const mobilePhone = escapeHTML(profile.mobilePhone);

  return `
<br />
  <table
    cellpadding="0"
    cellspacing="0"
    border="0"
    style="font-family: Arial, sans-serif; font-size: 10pt; line-height: 1.4"
  >
    <tr>
      <td style="padding-right: 12px; vertical-align: top">
        <img
          src="https://s3.us-east-1.amazonaws.com/intelligencehealthgroup.com/images/email_logo.png"
          width="165"
          alt="Intelligence Health Group"
          style="display: block; border: 0"
        />
      </td>
      <td
        style="
          border-left: 1px solid #999999;
          padding-left: 12px;
          vertical-align: top;
        "
      >
        <span style="color: #5b8fd6">${displayName} ${jobTitle ? `, ${jobTitle}` : ""}</span>
        <br /><br />
        ${mobilePhone ? `<b><span style="color: #2e6fd8">Cell:</span></b> <b>${mobilePhone}</b><br />` : ""}
        <b><span style="color: #2e6fd8">Office:</span></b> <b>(352) 421-5681 ${businessPhone}</b><br />
        <b><span style="color: #2e6fd8">Fax:</span></b> <b>(844) 927-4812</b>
        <br />
        <b><span style="color: #2e6fd8">Clinical On Call:</span></b>
        <b>(877) 294-4736</b><br /><br />
        <a
          href="https://maps.google.com/?q=1415+SW+17th+Street+Ocala+FL+34471"
          style="color: #333333; text-decoration: none"
          >1415 SW 17<sup>th</sup> Street<br />
          Ocala, FL 34471</a
        >
      </td>
    </tr>
  </table>
  <br />
`;
}

// Register the function with Office.
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
