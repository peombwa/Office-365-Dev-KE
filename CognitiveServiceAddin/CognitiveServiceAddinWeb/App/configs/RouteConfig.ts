module CognitiveServiceAddin.Configs {
    export class RouteConfig {
        constructor(private $locationProvider: ng.ILocationProvider,
            private $routeProvider: ng.route.IRouteProvider) {
            var that: RouteConfig = this;

            that.$routeProvider
                .when("/", {
                    controller: "TranslatorController",
                    controllerAs: "TranslatorCtrl",
                    templateUrl: "app/views/TranslatorPage.html"
                })

                .otherwise("/");

            that.$locationProvider.html5Mode(true);
        }
    }
}