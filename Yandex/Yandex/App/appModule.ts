module YandexAddin {

    export class App {
        constructor() {
            angular.module("CognitiveServiceApp", ["ngRoute"])
                .config(["$locationProvider", "$routeProvider", Configs.RouteConfig])
                .service("OfficeService", ["$log", Services.OfficeService])
                .service("TranslatorService", ["$log","$http",Services.TranslatorService])
                .controller("TranslatorController", ["$scope", "$log", "$timeout", "OfficeService","TranslatorService", Controllers.TranslatorController]);

            // initialize office.js lib
            Office.initialize = function(reason) {
                // when document is loaded, bootstrap our angular app
                angular.element(document).ready(() => {
                    angular.bootstrap(document, ["CognitiveServiceApp"]);
                });
            };

        }
    }

}