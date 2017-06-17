var YandexAddin;
(function (YandexAddin) {
    var App = (function () {
        function App() {
            angular.module("CognitiveServiceApp", ["ngRoute"])
                .config(["$locationProvider", "$routeProvider", YandexAddin.Configs.RouteConfig])
                .service("OfficeService", ["$log", YandexAddin.Services.OfficeService])
                .service("TranslatorService", ["$log", "$http", YandexAddin.Services.TranslatorService])
                .controller("TranslatorController", ["$scope", "$log", "$timeout", "OfficeService", "TranslatorService", YandexAddin.Controllers.TranslatorController]);
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
    YandexAddin.App = App;
})(YandexAddin || (YandexAddin = {}));
var YandexAddin;
(function (YandexAddin) {
    var Configs;
    (function (Configs) {
        var AppConfig = (function () {
            function AppConfig() {
            }
            AppConfig.TranslatorKey = "trnsl.1.1.20170616T132444Z.8a6358f25ffa04c9.4c7c24e4236de49256686cac785724cdc1e9f725";
            AppConfig.TranslatorEndpoint = "https://translate.yandex.net/api/v1.5/tr.json/translate";
            return AppConfig;
        }());
        Configs.AppConfig = AppConfig;
    })(Configs = YandexAddin.Configs || (YandexAddin.Configs = {}));
})(YandexAddin || (YandexAddin = {}));
var YandexAddin;
(function (YandexAddin) {
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
    })(Configs = YandexAddin.Configs || (YandexAddin.Configs = {}));
})(YandexAddin || (YandexAddin = {}));
var YandexAddin;
(function (YandexAddin) {
    var Controllers;
    (function (Controllers) {
        var TranslatorController = (function () {
            function TranslatorController($scope, $log, $timeout, officeService, translatorService) {
                this.$scope = $scope;
                this.$log = $log;
                this.$timeout = $timeout;
                this.officeService = officeService;
                this.translatorService = translatorService;
                var that = this;
            }
            TranslatorController.prototype.getBodyText = function () {
                var that = this;
                that.officeService.getBodyText()
                    .done(function (result) {
                    that.$timeout(0).then(function () {
                        that.$scope.selectedText = result;
                        console.log(result);
                    });
                })
                    .fail(function (error) {
                    console.error(error);
                });
            };
            TranslatorController.prototype.translateBodyText = function () {
                var that = this;
                that.translatorService.translateText(that.$scope.selectedText, "en", "ru")
                    .done(function (result) {
                    that.$timeout(0).then(function () {
                        that.$scope.translatedText = result;
                    });
                })
                    .fail(function (error) {
                    console.error(error);
                });
            };
            return TranslatorController;
        }());
        Controllers.TranslatorController = TranslatorController;
    })(Controllers = YandexAddin.Controllers || (YandexAddin.Controllers = {}));
})(YandexAddin || (YandexAddin = {}));
var YandexAddin;
(function (YandexAddin) {
    var Services;
    (function (Services) {
        var OfficeService = (function () {
            function OfficeService($log) {
                this.$log = $log;
                this.mailBoxItem = null;
                var that = this;
                that.mailBoxItem = Office.context.mailbox.item;
            }
            OfficeService.prototype.getBodyText = function () {
                var that = this;
                var deferred = $.Deferred();
                that.mailBoxItem.body.getAsync("text", function (result) {
                    if (result.status == "succeeded") {
                        deferred.resolve(result.value);
                    }
                    else {
                        deferred.reject(result);
                    }
                });
                return deferred;
            };
            OfficeService.prototype.writeTextToMail = function (text) {
                var that = this;
                var deferred = $.Deferred();
                that.mailBoxItem.body.prependAsync(text, { coercionType: Office.CoercionType.Text, asyncContext: { var3: 1, var4: 2 } }, function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        that.$log.error("Failed to add text to the body: ", asyncResult.error.message);
                        deferred.reject(asyncResult.error.message);
                    }
                    else {
                        that.$log.debug("Added text to the body: ", that.mailBoxItem.body);
                        deferred.resolve(true);
                    }
                });
                return deferred;
            };
            return OfficeService;
        }());
        Services.OfficeService = OfficeService;
    })(Services = YandexAddin.Services || (YandexAddin.Services = {}));
})(YandexAddin || (YandexAddin = {}));
var YandexAddin;
(function (YandexAddin) {
    var Services;
    (function (Services) {
        var TranslatorService = (function () {
            function TranslatorService($log, $http) {
                this.$log = $log;
                this.$http = $http;
                var that = this;
            }
            TranslatorService.prototype.translateText = function (text, to, from) {
                var that = this;
                var deferred = $.Deferred();
                var qureryString = "?key=" + YandexAddin.Configs.AppConfig.TranslatorKey + "&text=" + text + "&lang=" + from + "-" + to + "&format=plain";
                // issue request to translate
                that.$http({
                    method: 'GET',
                    url: YandexAddin.Configs.AppConfig.TranslatorEndpoint + qureryString,
                    headers: {
                        'Content-Type': 'application/json'
                    }
                }).then(function (success) {
                    that.$log.debug("Successfully translated text: ", success);
                    deferred.resolve(success.data.text[0]);
                }, function (error) {
                    that.$log.error("Failed to translate text: ", error);
                    deferred.reject(error);
                });
                return deferred;
            };
            return TranslatorService;
        }());
        Services.TranslatorService = TranslatorService;
    })(Services = YandexAddin.Services || (YandexAddin.Services = {}));
})(YandexAddin || (YandexAddin = {}));
var YandexAddin;
(function (YandexAddin) {
    var Services;
    (function (Services) {
        var WordService = (function () {
            function WordService($log) {
                this.$log = $log;
                var that = this;
            }
            WordService.prototype.getBodyText = function () {
                var that = this;
                var deferred = $.Deferred();
                Word.run(function (ctx) {
                    // create proxy object for the document
                    var docBody = ctx.document.body;
                    // queue command to load the text property of proxy object
                    ctx.load(docBody, "text");
                    // sync the doument object with the proxy objects
                    return ctx.sync().then(function () {
                        that.$log.debug("Body contents: ", docBody.text);
                        deferred.resolve(docBody.text);
                    }).catch(function (error) {
                        that.$log.error("Failed to get text from the body: ", error);
                        deferred.reject(error);
                    });
                });
                return deferred;
            };
            WordService.prototype.writeTextToBody = function (text) {
                var that = this;
                var deferred = $.Deferred();
                Word.run(function (ctx) {
                    var docBody = ctx.document.body;
                    ctx.load(docBody, "text");
                    docBody.insertText(text, Word.InsertLocation.end);
                    return ctx.sync().then(function () {
                        that.$log.debug("New body contents: ", docBody.text);
                        deferred.resolve(true);
                    }).catch(function (error) {
                        that.$log.error("Failed to add text to the body: ", error);
                        deferred.reject(error);
                    });
                });
                return deferred;
            };
            return WordService;
        }());
        Services.WordService = WordService;
    })(Services = YandexAddin.Services || (YandexAddin.Services = {}));
})(YandexAddin || (YandexAddin = {}));
//# sourceMappingURL=app.js.map