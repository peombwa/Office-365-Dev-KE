var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var App = (function () {
        function App() {
            angular.module("CognitiveServiceApp", ["ngRoute"])
                .config(["$locationProvider", "$routeProvider", CognitiveServiceAddin.Configs.RouteConfig])
                .service("OfficeService", ["$log", CognitiveServiceAddin.Services.OfficeService])
                .controller("MainController", ["$scope", "$log", CognitiveServiceAddin.Controllers.MainController])
                .controller("TranslatorController", ["$scope", "$log", "OfficeService", CognitiveServiceAddin.Controllers.TranslatorController]);
            // initialize office.js lib
            Office.initialize = function (reason) {
                // when document is loaded, bootstrap our angular app
                angular.element(document).ready(function () {
                    angular.bootstrap(document, ["CognitiveServiceApp"]);
                });
            };
        }
        return App;
    }());
    CognitiveServiceAddin.App = App;
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Configs;
    (function (Configs) {
        var RouteConfig = (function () {
            function RouteConfig($locationProvider, $routeProvider) {
                this.$locationProvider = $locationProvider;
                this.$routeProvider = $routeProvider;
                var that = this;
                that.$routeProvider
                    .when("/", {
                    controller: "TranslatorController",
                    controllerAs: "TranslatorCtrl",
                    templateUrl: "app/views/TranslatorPage.html"
                })
                    .otherwise("/");
                that.$locationProvider.html5Mode(true);
            }
            return RouteConfig;
        }());
        Configs.RouteConfig = RouteConfig;
    })(Configs = CognitiveServiceAddin.Configs || (CognitiveServiceAddin.Configs = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Controllers;
    (function (Controllers) {
        var MainController = (function () {
            function MainController($scope, $log) {
                this.$scope = $scope;
                this.$log = $log;
                var that = this;
            }
            return MainController;
        }());
        Controllers.MainController = MainController;
    })(Controllers = CognitiveServiceAddin.Controllers || (CognitiveServiceAddin.Controllers = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Controllers;
    (function (Controllers) {
        var TranslatorController = (function () {
            function TranslatorController($scope, $log, OfficeService) {
                this.$scope = $scope;
                this.$log = $log;
                this.OfficeService = OfficeService;
                var that = this;
            }
            return TranslatorController;
        }());
        Controllers.TranslatorController = TranslatorController;
    })(Controllers = CognitiveServiceAddin.Controllers || (CognitiveServiceAddin.Controllers = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Services;
    (function (Services) {
        var OfficeService = (function () {
            function OfficeService($log) {
                this.$log = $log;
                var that = this;
            }
            OfficeService.prototype.getSelectedText = function () {
                var that = this;
                return "";
            };
            OfficeService.prototype.writeTextToMail = function (text) {
                var that = this;
            };
            return OfficeService;
        }());
        Services.OfficeService = OfficeService;
    })(Services = CognitiveServiceAddin.Services || (CognitiveServiceAddin.Services = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
//# sourceMappingURL=app.js.map