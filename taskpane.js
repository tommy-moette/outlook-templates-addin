// Configuration - Update these values
const msalConfig = {
    auth: {
        clientId: "ca529e3c-88b4-4db3-aacb-514f2882b081", // Register app in Azure AD
        authority: "https://login.microsoftonline.com/c0ae7a25-3b55-401b-8a09-f431d96e686f",
        redirectUri: "https://gray-desert-02be92c03.3.azurestaticapps.net/taskpane.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

//const loginRequest = {
//    scopes: ["Files.Read.All", "Sites.Read.All"]
//};

const loginRequest = {
    scopes: ["https://mo1e.sharepoint.com/.default"]
};


// SharePoint configuration
const SHAREPOINT_SITE = "https://mo1e.sharepoint.com/sites/EmailTemplates";
const TEMPLATES_FOLDER = "/Delade dokument/EmailTemplates";

let msalInstance;
let accessToken = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        
        document.getElementById("login-btn").onclick = signIn;
        
        // Check if already signed in
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
    if (!account) {
        throw new Error("No active account");
    }
    
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
console.log("Access token:", accessToken);

async function loadTemplates() {
    const templatesList = document.getElementById("templates-list");
    const loading = document.getElementById("loading");
    
    templatesList.innerHTML = "";
    loading.style.display = "block";
    
    try {
        // Get files from SharePoint folder
        const siteUrl = `${SHAREPOINT_SITE}/_api/web/GetFolderByServerRelativeUrl('${TEMPLATES_FOLDER}')/Files`;
        
        const response = await fetch(siteUrl, {
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Accept": "application/json;odata=verbose"
            }
        });
        
        if (!response.ok) {
            throw new Error("Failed to fetch templates");
        }
        
        const data = await response.json();
        const templates = data.d.results.filter(file => 
            file.Name.endsWith('.html') || file.Name.endsWith('.htm')
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
                <div class="template-name">${template.Name.replace(/\.(html|htm)$/, '')}</div>
                <div class="template-desc">Click to insert</div>
            `;
            li.onclick = () => insertTemplate(template.ServerRelativeUrl);
            templatesList.appendChild(li);
        });
        
    } catch (error) {
        loading.style.display = "none";
        showStatus("Error loading templates: " + error.message, "error");
    }
}

async function insertTemplate(fileUrl) {
    try {
        // Fetch template content
        const downloadUrl = `${SHAREPOINT_SITE}/_api/web/GetFileByServerRelativeUrl('${fileUrl}')/$value`;
        
        const response = await fetch(downloadUrl, {
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });
        
        if (!response.ok) {
            throw new Error("Failed to fetch template content");
        }
        
        const htmlContent = await response.text();
        
        // Insert into email
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
