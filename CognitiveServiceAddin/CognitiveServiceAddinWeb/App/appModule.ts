module CognitiveServiceAddin {

    export class App {
        constructor() {
            angular.module("CognitiveServiceApp", ["ngRoute"])
                .controller("MainController", ["$scope","$log",Controllers.MainController]);

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