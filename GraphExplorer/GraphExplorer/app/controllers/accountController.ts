module TimeFinder.Controllers {
    export class AccountController {

        constructor(private $log: ng.ILogService,
            private $timeout: ng.ITimeoutService,
            private $location:ng.ILocationService,
            private authService: Services.AuthService) {

            var that: AccountController = this;
            that.init();
        }

        private init() {
            var that: AccountController = this;

            if (that.authService.isUserLoggedIn('aad') && !that.authService.isTokenExpired('aad')) {
                // user has logged in and the token has not expired or is not about to expire, proceed with the application
                that.$location.path('main');
            }
            else if (that.authService.isUserLoggedIn('aad') && that.authService.isTokenExpired('aad')) {
                // user has logged in and the token has expired or is about to expire, refresh the token
                that.authService.refreshToken('aad')
                    .done(() => {
                        that.$timeout(0)
                            .then(() => {
                                that.$location.path('main');
                            });
                    })
                    .fail(() => {

                    });

            }
        }

        public signInWithAD() {
            //Login User
            var that: AccountController = this;

            that.$log.debug("Logging in..");
            that.authService.login('aad')
                .done(() => {
                    that.$log.debug("Login success");
                    that.$timeout(0)
                        .then(() => {
                            that.$location.path('main');
                        });
                })
                .fail((error) => {
                    that.$log.error("User did not login.", error);
                });

        }
    }
}