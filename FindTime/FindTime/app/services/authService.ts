
module TimeFinder.Services {
    

    interface IAuthResponse {
        access_token?: string;
        display?: string;
        expires?: number;
        expires_in?: number;
        network?: string;
        redirect_uri?: string;
        scope?: string;
        session_state?: string;
        state?: string;
        token_type?: string;
    }

    export class AuthService {

        private loggedInUser: Models.IUser = {};

        constructor(private $log: ng.ILogService) {
            var that: AuthService = this;

            hello.on('auth.login', function (r) {

                // Get Profile
                hello("aad").api('me').then(function (p: Models.IUser) {
                    that.loggedInUser = p;

                    that.$log.debug("User is : ", p);
                });
            });

        }

        private log(response) {
            var that: AuthService = this;

            that.$log.warn("Response", response);
        }

        public getLoggedInUser(): Models.IUser {
            var that: AuthService = this;

            return that.loggedInUser;
        }

        public login(network: string): JQueryDeferred<any> {
            var deferred = $.Deferred();
            var that: AuthService = this;
            // By defining response type to code, the OAuth flow that will return a refresh token to be used to refresh the access token
            // However this will require the oauth_proxy server
            hello(network).login({ display: 'popup' }, that.log)
                .then((response) => {

                        // Get Profile
                    hello("aad").api('me').then(function (p: Models.IUser) {
                            that.loggedInUser = p;

                            that.$log.debug("User is : ", p);
                            that.$log.debug('You are signed in to AAD', response);

                            deferred.resolve(response);
                        });
                    },
                    (e) => {
                        that.$log.error('Signin error: ' + e.error.message);
                        deferred.reject(e);
                    });

            return deferred;
        }

        public logout(network): JQueryDeferred<any> {
            var deferred = $.Deferred();
            var that: AuthService = this;

            // Removes all sessions, need to call AAD endpoint to do full logout
            hello(network).logout({ force: true }, () => {}).then(function () {
                that.$log.debug('You have Signed Out of  AAD');

                deferred.resolve();

            }, function (e) {
                that.$log.error('Sign out error: ' + e.error.message);

                deferred.reject(e);
            });

            return deferred;
        }

        public getToken(network): JQueryDeferred<string> {
            var deferred = $.Deferred();
            var that: AuthService = this;

            if (that.isUserLoggedIn(network) && !that.isTokenExpired(network)) {
                var authResponse: IAuthResponse = hello.getAuthResponse('aad');

                deferred.resolve(authResponse.access_token);
            } else {
                that.login(network)
                    .done(() => {
                        var authResponse: IAuthResponse = hello.getAuthResponse('aad');

                        deferred.resolve(authResponse.access_token);
                    })
                    .fail((error) => {
                        deferred.reject(null);
                    });
            }

            return deferred;
        }

        public refreshToken(network): JQueryDeferred<any> {
            var deferred = $.Deferred();
            var that: AuthService = this;

            hello(network).login({ display: 'popup', force: false }, that.log).then(function () {

                hello("aad").api('me').then(function (p: Models.IUser) {
                    that.loggedInUser = p;

                    that.$log.debug("User is : ", p);

                    that.$log.debug('Token refreshed....');

                    deferred.resolve();
                });



            }, function (e) {
                that.$log.error('Token refresh error: ' + e.error.message);

                deferred.reject(e);
            });

            return deferred;
        }

        public isTokenExpired(network): boolean {
            var that: AuthService = this;

            var authResponse: IAuthResponse = hello.getAuthResponse('aad');

            var expiryDate = moment.unix(authResponse.expires);
            var duration = moment.duration(expiryDate.diff(moment.now())).asMinutes();

            that.$log.debug("Token expires in : ", duration);

            // less than 10 minutes
            var isTokenExpired = ((duration < 10) ? true : false);
            that.$log.warn("Token expired: ", isTokenExpired);

            return isTokenExpired;
        }

        public isUserLoggedIn(netrowk): boolean {
            var that: AuthService = this;

            var authResponse: IAuthResponse = hello.getAuthResponse('aad');
            that.$log.warn("User logged in auth response : ", authResponse);


            if (!_.isNull(authResponse)) {
                // user has logged in to the app before
                // check if the loggin was successfull -- uses expires property
                var response = (_.has(authResponse, "expires") ? true : false);
                that.$log.warn("User logged in? : ", response);

                return response;

            } else {
                // user hasn't used the app so the local storage is empty
                that.$log.error("User logged in auth response is null : ", authResponse);
                return false;
            }
        }

    }
}