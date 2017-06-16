module TimeFinder {
    export class App {
        constructor() {
            angular.module("TimeFinder", ["ngRoute"])
                .config(["$locationProvider", "$routeProvider", "$httpProvider", Configs.RouteConfig])
                .config([Configs.AuthConfig])                       
                .service("AuthService", ["$log", Services.AuthService])
                .service("GraphService", ["$log", "$http","AuthService", Services.GraphService])
                .controller("AccountController", ["$log", "$timeout", "$location","AuthService", Controllers.AccountController])
                .controller("MainController", ["$log", "$scope","$timeout","GraphService", Controllers.MainController]);

            angular.bootstrap(document, ["TimeFinder"]);

        }
    }
}