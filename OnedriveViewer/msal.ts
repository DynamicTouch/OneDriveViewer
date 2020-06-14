import { UserAgentApplication, AuthenticationParameters } from 'msal';

export interface IAquireTokenProperties {
    validateAuthority: boolean;
    clientId: string;
    authority: string;
    redirectUri: string;
    scopes: string[];
}


export class MSALConnection {


    private isInteractionRequired(error: Error): boolean {
        if (error.message?.length <= 0) {
            return false;
        }

        return (
            error.message.indexOf('consent_required') > -1 ||
            error.message.indexOf('interaction_required') > -1 ||
            error.message.indexOf('login_required') > -1
        );
    }

    public async AquireToken(properties: IAquireTokenProperties): Promise<string> {

        const msalConfig = {
            auth: {
                clientId: properties.clientId,
                authority: properties.authority,
                validateAuthority: properties.validateAuthority,
            }
        };

        const msalInstance = new UserAgentApplication(msalConfig);

        const tokenRequest: AuthenticationParameters = {
            scopes: properties.scopes,
            redirectUri: properties.redirectUri
        };

        if (msalInstance.getAccount()) {
            return await msalInstance.acquireTokenSilent(tokenRequest)
                .then(response => {
                    return response.accessToken;
                })
                .catch(err => {
                    if (this.isInteractionRequired(err)) {
                        return msalInstance.acquireTokenPopup(tokenRequest)
                            .then(response => {
                                return response.accessToken;
                            })
                            .catch(err => {
                                console.log(err);
                                alert(err.message);
                                return '';
                            });
                    }
                    return '';
                });
        }
        else {
            const loginRequest = {
                scopes: properties.scopes,
            };
            return await msalInstance.loginPopup(loginRequest)
                .then(response => {
                    return response.accessToken;
                }
                )
                .catch(err => {
                    console.log('login popup error ' + err);
                    alert(err.message);
                    return '';
                });
        }
        return '';
    }


}