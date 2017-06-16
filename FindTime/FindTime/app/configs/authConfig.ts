module TimeFinder.Configs {
    export class AuthConfig {
        constructor() {
            hello.init({
                aad: `${Configs.AppConfig.ClientId}`
            }, {
                redirect_uri: '../',
                scope: 'openid email profile Calendars.ReadWrite.Shared Calendars.ReadWrite Contacts.Read User.ReadBasic.All'
            });
        }
    }
}