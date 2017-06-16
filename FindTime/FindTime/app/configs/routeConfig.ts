module TimeFinder.Configs {
    //interface IAdalRouteProvider extends ng.route.IRouteProvide {
    //    requireADLogin?:boolean;
    //}
    export class RouteConfig {
        constructor(private $locationProvider: ng.ILocationProvider,
            private $routeProvider: ng.route.IRouteProvider,
            private $httpProvider: ng.IHttpProvider) {
            var that: RouteConfig = this;

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
    }
}