export default class MsalAuthenticator {
    constructor() {
        const msalConfig = {
            auth: {
                authority: 'https://login.microsoftonline.com/organizations/',
                clientId: '9c4e11f7-0b9f-48e8-8e57-3809160039aa',
                redirectUri: 'http://localhost:5146/',
            },
            cache: {
                cacheLocation: 'localStorage'
            }
        };

        this.powerAppsTokenRequest = {
            scopes: ['https://service.flow.microsoft.com/.default']
        };

        this.msalInstance = new msal.PublicClientApplication(msalConfig);
        this.msalInstance.initialize();
    }

    async authenticate() {
        const request = structuredClone(this.powerAppsTokenRequest);
        const currentAccount = localStorage.getItem('currentAccount');

        if (currentAccount) {
            request.account = JSON.parse(currentAccount);
        }

        var response = null;
        var scenario = 'Cache';

        try {
            response = await this.msalInstance.acquireTokenSilent(request);
        } catch (x) {
            console.debug(x);
            if (x instanceof msal.BrowserAuthError) {
                try {
                    response = await this.msalInstance.ssoSilent(request);
                    scenario = 'SSO';
                } catch (xx) {
                    console.debug(xx);
                    if (xx instanceof msal.InteractionRequiredAuthError) {
                        response = await this.msalInstance.acquireTokenPopup(powerAppsTokenRequest);
                        scenario = 'Interactive';
                    }
                }
            }
        }

        if (response) {
            const account = response.account;
            console.log(`${scenario} token acquisition successful:`, `${account.username} (${account.tenantId})`);
            localStorage.setItem('currentAccount', JSON.stringify(account));
        }

        return response;
    }
}