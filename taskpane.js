// Configuration
const msalConfig = {
    auth: {
        clientId: "ca529e3c-88b4-4db3-aacb-514f2882b081",
        authority: "https://login.microsoftonline.com/c0ae7a25-3b55-401b-8a09-f431d96e686f",
        redirectUri: "https://gray-desert-02be92c03.3.azurestaticapps.net/taskpane.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

const loginRequest = {
    scopes: ["Files.Read.All", "Sites.Read.All"]
};

let msalInstance;
let accessToken = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        document.getElementById("login-btn").onclick = signIn;
        checkAuthState();
    }
});

async function checkAuthState() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
        document.getElementById("login-btn").style.display = "none";
        await getAccessToken();
        await loadTemplates();
    }
}

async function signIn() {
    try {
        const loginResponse = await msalInstance.loginPopup(loginRequest);
        msalInstance.setActiveAccount(loginResponse.account);
        document.getElementById("login-btn").style.display = "none";
        await getAccessToken();
        await loadTemplates();
    } catch (error) {
        showStatus("Sign-in failed: " + error.message, "error");
    }
}

async function getAccessToken() {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("No active account");

    try {
        const response = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: account
        });
        accessToken = response.accessToken;
    } catch (error) {
        const response = await msalInstance.acquireTokenPopup(loginRequest);
        accessToken = response.accessToken;
    }
}

async function loadTemplates() {
    const templatesList = document.getElementById("templates-list");
    const loading = document.getElementById("loading");

    templatesList.innerHTML = "";
    loading.style.display = "block";

    try {
        // 1. Get site ID
        const siteResponse = await fetch("https://graph.microsoft.com/v1.0/sites/mo1e.sharepoint.com:/sites/EmailTemplates", {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        if (!siteResponse.ok) {
            const errorText = await siteResponse.text();
            throw new Error(`Failed to get site ID: ${siteResponse.status} ${siteResponse.statusText} - ${errorText}`);
        }

        const siteData = await siteResponse.json();
        const siteId = siteData.id;

        // 2. Get files from folder
        const folderPath = "01ZTYGZROGB3NR3JVFEFELH72IFM4ZGSJU";
        const filesResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${folderPath}:/children`, {
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        if (!filesResponse.ok) {
            const errorText = await filesResponse.text();
            throw new Error(`Failed to get files: ${filesResponse.status} ${filesResponse.statusText} - ${errorText}`);
        }

        const filesData = await filesResponse.json();
        const templates = filesData.value.filter(file =>
            file.name.endsWith('.html') || file.name.endsWith('.htm')
        );

        loading.style.display = "none";

        if (templates.length === 0) {
            templatesList.innerHTML = '<li style="padding: 20px; text-align: center; color: #605e5c;">No templates found</li>';
            return;
        }

        templates.forEach(template => {
            const li = document.createElement("li");
            li.className = "template-item";
            li.innerHTML = `
                <div class="template-name">${template.name.replace(/\.(html|htm)$/, '')}</div>
                <div class="template-desc">Click to insert</div>
            `;
            li.onclick = () => insertTemplate(template['@microsoft.graph.downloadUrl']);
            templatesList.appendChild(li);
        });

    } catch (error) {
        loading.style.display = "none";
        showStatus("Error loading templates: " + error.message, "error");
    }
}

async function insertTemplate(downloadUrl) {
    try {
        const response = await fetch(downloadUrl);
        if (!response.ok) {
            throw new Error("Failed to fetch template content");
        }

        const htmlContent = await response.text();

        Office.context.mailbox.item.body.setAsync(
            htmlContent,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showStatus("✓ Template inserted successfully!", "success");
                } else {
                    showStatus("Error inserting template: " + result.error.message, "error");
                }
            }
        );
    } catch (error) {
        showStatus("Error: " + error.message, "error");
    }
}

function showStatus(message, type) {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = type;
    status.style.display = "block";

    setTimeout(() => {
        status.style.display = "none";
    }, 4000);
}
