var TimeFinder;
(function (TimeFinder) {
    var App = (function () {
        function App() {
            angular.module("TimeFinder", ["ngRoute"])
                .config(["$locationProvider", "$routeProvider", "$httpProvider", TimeFinder.Configs.RouteConfig])
                .config([TimeFinder.Configs.AuthConfig])
                .service("AuthService", ["$log", TimeFinder.Services.AuthService])
                .service("GraphService", ["$log", "$http", "AuthService", TimeFinder.Services.GraphService])
                .controller("AccountController", ["$log", "$timeout", "$location", "AuthService", TimeFinder.Controllers.AccountController])
                .controller("MainController", ["$log", "$scope", "$timeout", "GraphService", TimeFinder.Controllers.MainController]);
            angular.bootstrap(document, ["TimeFinder"]);
        }
        return App;
    }());
    TimeFinder.App = App;
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Configs;
    (function (Configs) {
        var AppConfig = (function () {
            function AppConfig() {
            }
            AppConfig.GraphEndPoint = "https://graph.microsoft.com/v1.0";
            AppConfig.ClientId = "2c82f2b7-9179-4f9a-83f8-dc1b452fc5a2";
            return AppConfig;
        }());
        Configs.AppConfig = AppConfig;
    })(Configs = TimeFinder.Configs || (TimeFinder.Configs = {}));
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Configs;
    (function (Configs) {
        var AuthConfig = (function () {
            function AuthConfig() {
                hello.init({
                    aad: "" + Configs.AppConfig.ClientId
                }, {
                    redirect_uri: '../',
                    scope: 'openid email profile Calendars.ReadWrite.Shared Calendars.ReadWrite Contacts.Read User.ReadBasic.All'
                });
            }
            return AuthConfig;
        }());
        Configs.AuthConfig = AuthConfig;
    })(Configs = TimeFinder.Configs || (TimeFinder.Configs = {}));
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Configs;
    (function (Configs) {
        //interface IAdalRouteProvider extends ng.route.IRouteProvide {
        //    requireADLogin?:boolean;
        //}
        var RouteConfig = (function () {
            function RouteConfig($locationProvider, $routeProvider, $httpProvider) {
                this.$locationProvider = $locationProvider;
                this.$routeProvider = $routeProvider;
                this.$httpProvider = $httpProvider;
                var that = this;
                that.$routeProvider
                    .when("/", {
                    controller: "AccountController",
                    controllerAs: "AccountCtrl",
                    templateUrl: "app/views/account.html"
                })
                    .when("/main", {
                    controller: "MainController",
                    controllerAs: "MainCtrl",
                    templateUrl: "app/views/home.html"
                })
                    .otherwise("/");
                that.$locationProvider.html5Mode(true).hashPrefix('!');
            }
            return RouteConfig;
        }());
        Configs.RouteConfig = RouteConfig;
    })(Configs = TimeFinder.Configs || (TimeFinder.Configs = {}));
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Controllers;
    (function (Controllers) {
        var AccountController = (function () {
            function AccountController($log, $timeout, $location, authService) {
                this.$log = $log;
                this.$timeout = $timeout;
                this.$location = $location;
                this.authService = authService;
                var that = this;
                that.init();
            }
            AccountController.prototype.init = function () {
                var that = this;
                if (that.authService.isUserLoggedIn('aad') && !that.authService.isTokenExpired('aad')) {
                    // user has logged in and the token has not expired or is not about to expire, proceed with the application
                    that.$location.path('main');
                }
                else if (that.authService.isUserLoggedIn('aad') && that.authService.isTokenExpired('aad')) {
                    // user has logged in and the token has expired or is about to expire, refresh the token
                    that.authService.refreshToken('aad')
                        .done(function () {
                        that.$timeout(0)
                            .then(function () {
                            that.$location.path('main');
                        });
                    })
                        .fail(function () {
                    });
                }
            };
            AccountController.prototype.signInWithAD = function () {
                //Login User
                var that = this;
                that.$log.debug("Logging in..");
                that.authService.login('aad')
                    .done(function () {
                    that.$log.debug("Login success");
                    that.$timeout(0)
                        .then(function () {
                        that.$location.path('main');
                    });
                })
                    .fail(function (error) {
                    that.$log.error("User did not login.", error);
                });
            };
            return AccountController;
        }());
        Controllers.AccountController = AccountController;
    })(Controllers = TimeFinder.Controllers || (TimeFinder.Controllers = {}));
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Controllers;
    (function (Controllers) {
        var MainController = (function () {
            function MainController($log, $scope, $timeout, graphService) {
                this.$log = $log;
                this.$scope = $scope;
                this.$timeout = $timeout;
                this.graphService = graphService;
                var that = this;
            }
            MainController.prototype.findMeetingTimes = function () {
                var that = this;
                var attendees = [];
                angular.forEach(that.$scope.adUsers, function (value, key) {
                    attendees.push({
                        type: "required",
                        emailAddress: {
                            address: value.mail,
                            name: value.displayName
                        }
                    });
                });
                // set time constraiint
                var timeConstraint = {
                    activityDomain: "unrestricted",
                    timeslots: [
                        {
                            start: {
                                //2017-04-17T09:00:00
                                dateTime: moment().format(),
                                timeZone: "UTC"
                            },
                            end: {
                                dateTime: moment().add(5, "hours").format(),
                                timeZone: "UTC"
                            }
                        }]
                };
                that.graphService.findMeetingTimes({ attendees: attendees, timeConstraint: timeConstraint, minimumAttendeePercentage: 10, returnSuggestionReasons: true })
                    .done(function (response) {
                    that.$log.debug("Response: ", response);
                    that.$timeout(0).then(function () {
                        that.$scope.meetingTimeSuggestions = response.meetingTimeSuggestions;
                    });
                })
                    .fail(function (error) {
                    that.$log.error("Error: ", error);
                });
            };
            MainController.prototype.listADUsers = function () {
                var that = this;
                that.graphService.listUsers()
                    .done(function (response) {
                    that.$log.debug("Users: ", response);
                    that.$timeout(0).then(function () {
                        that.$scope.adUsers = response;
                    });
                })
                    .fail(function (error) {
                    that.$log.error("Error: ", error);
                });
            };
            return MainController;
        }());
        Controllers.MainController = MainController;
    })(Controllers = TimeFinder.Controllers || (TimeFinder.Controllers = {}));
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Services;
    (function (Services) {
        var AuthService = (function () {
            function AuthService($log) {
                this.$log = $log;
                this.loggedInUser = {};
                var that = this;
                hello.on('auth.login', function (r) {
                    // Get Profile
                    hello("aad").api('me').then(function (p) {
                        that.loggedInUser = p;
                        that.$log.debug("User is : ", p);
                    });
                });
            }
            AuthService.prototype.log = function (response) {
                var that = this;
                that.$log.warn("Response", response);
            };
            AuthService.prototype.getLoggedInUser = function () {
                var that = this;
                return that.loggedInUser;
            };
            AuthService.prototype.login = function (network) {
                var deferred = $.Deferred();
                var that = this;
                // By defining response type to code, the OAuth flow that will return a refresh token to be used to refresh the access token
                // However this will require the oauth_proxy server
                hello(network).login({ display: 'popup' }, that.log)
                    .then(function (response) {
                    // Get Profile
                    hello("aad").api('me').then(function (p) {
                        that.loggedInUser = p;
                        that.$log.debug("User is : ", p);
                        that.$log.debug('You are signed in to AAD', response);
                        deferred.resolve(response);
                    });
                }, function (e) {
                    that.$log.error('Signin error: ' + e.error.message);
                    deferred.reject(e);
                });
                return deferred;
            };
            AuthService.prototype.logout = function (network) {
                var deferred = $.Deferred();
                var that = this;
                // Removes all sessions, need to call AAD endpoint to do full logout
                hello(network).logout({ force: true }, function () { }).then(function () {
                    that.$log.debug('You have Signed Out of  AAD');
                    deferred.resolve();
                }, function (e) {
                    that.$log.error('Sign out error: ' + e.error.message);
                    deferred.reject(e);
                });
                return deferred;
            };
            AuthService.prototype.getToken = function (network) {
                var deferred = $.Deferred();
                var that = this;
                if (that.isUserLoggedIn(network) && !that.isTokenExpired(network)) {
                    var authResponse = hello.getAuthResponse('aad');
                    deferred.resolve(authResponse.access_token);
                }
                else {
                    that.login(network)
                        .done(function () {
                        var authResponse = hello.getAuthResponse('aad');
                        deferred.resolve(authResponse.access_token);
                    })
                        .fail(function (error) {
                        deferred.reject(null);
                    });
                }
                return deferred;
            };
            AuthService.prototype.refreshToken = function (network) {
                var deferred = $.Deferred();
                var that = this;
                hello(network).login({ display: 'popup', force: false }, that.log).then(function () {
                    hello("aad").api('me').then(function (p) {
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
            };
            AuthService.prototype.isTokenExpired = function (network) {
                var that = this;
                var authResponse = hello.getAuthResponse('aad');
                var expiryDate = moment.unix(authResponse.expires);
                var duration = moment.duration(expiryDate.diff(moment.now())).asMinutes();
                that.$log.debug("Token expires in : ", duration);
                // less than 10 minutes
                var isTokenExpired = ((duration < 10) ? true : false);
                that.$log.warn("Token expired: ", isTokenExpired);
                return isTokenExpired;
            };
            AuthService.prototype.isUserLoggedIn = function (netrowk) {
                var that = this;
                var authResponse = hello.getAuthResponse('aad');
                that.$log.warn("User logged in auth response : ", authResponse);
                if (!_.isNull(authResponse)) {
                    // user has logged in to the app before
                    // check if the loggin was successfull -- uses expires property
                    var response = (_.has(authResponse, "expires") ? true : false);
                    that.$log.warn("User logged in? : ", response);
                    return response;
                }
                else {
                    // user hasn't used the app so the local storage is empty
                    that.$log.error("User logged in auth response is null : ", authResponse);
                    return false;
                }
            };
            return AuthService;
        }());
        Services.AuthService = AuthService;
    })(Services = TimeFinder.Services || (TimeFinder.Services = {}));
})(TimeFinder || (TimeFinder = {}));
var TimeFinder;
(function (TimeFinder) {
    var Services;
    (function (Services) {
        var GraphService = (function () {
            function GraphService($log, $http, authService) {
                this.$log = $log;
                this.$http = $http;
                this.authService = authService;
                var that = this;
            }
            GraphService.prototype.findMeetingTimes = function (suggestionRequest) {
                var that = this;
                var deferred = $.Deferred();
                // get cached token
                that.authService.getToken('aad')
                    .done(function (token) {
                    // make http post request
                    that.$http({
                        method: 'POST',
                        url: TimeFinder.Configs.AppConfig.GraphEndPoint + '/me/findMeetingTimes',
                        data: JSON.stringify(suggestionRequest),
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + token
                        }
                    })
                        .then(function (successresponse) {
                        deferred.resolve(successresponse.data);
                    }, function (errorResponse) {
                        deferred.reject(errorResponse);
                    });
                })
                    .fail(function (error) {
                    that.$log.error("Failed to fetch token");
                    deferred.reject(null);
                });
                return deferred;
            };
            GraphService.prototype.listUsers = function () {
                var that = this;
                var deferred = $.Deferred();
                that.authService.getToken('aad')
                    .done(function (token) {
                    that.$http({
                        method: 'GET',
                        url: TimeFinder.Configs.AppConfig.GraphEndPoint + '/users/',
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + token
                        }
                    })
                        .then(function (successresponse) {
                        that.$log.debug("successresponse: ", successresponse);
                        deferred.resolve(successresponse.data.value);
                    }, function (errorResponse) {
                        deferred.reject(errorResponse);
                    });
                });
                return deferred;
            };
            return GraphService;
        }());
        Services.GraphService = GraphService;
    })(Services = TimeFinder.Services || (TimeFinder.Services = {}));
})(TimeFinder || (TimeFinder = {}));
//# sourceMappingURL=app.js.map