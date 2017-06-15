var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var App = (function () {
        function App() {
            angular.module("CognitiveServiceApp", ["ngRoute"])
                .controller("MainController", ["$scope", "$log", CognitiveServiceAddin.Controllers.MainController]);
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
//# sourceMappingURL=app.js.map