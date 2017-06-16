var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var App = (function () {
        function App() {
            angular.module("CognitiveServiceApp", ["ngRoute"])
                .config(["$locationProvider", "$routeProvider", CognitiveServiceAddin.Configs.RouteConfig])
                .service("OfficeService", ["$log", CognitiveServiceAddin.Services.OfficeService])
                .service("TranslatorService", ["$log", "$http", CognitiveServiceAddin.Services.TranslatorService])
                .controller("MainController", ["$scope", "$log", CognitiveServiceAddin.Controllers.MainController])
                .controller("TranslatorController", ["$scope", "$log", "$timeout", "OfficeService", "TranslatorService", CognitiveServiceAddin.Controllers.TranslatorController]);
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
        var AppConfig = (function () {
            function AppConfig() {
            }
            AppConfig.AccessTokenUrl = "https://api.cognitive.microsoft.com/sts/v1.0/issueToken";
            AppConfig.TranslatorKey = "1514575dc1e643a6a511cf1e35d5819f";
            AppConfig.TranslatorEndpoint = "https://api.microsofttranslator.com/V2/Http.svc/Translate";
            return AppConfig;
        }());
        Configs.AppConfig = AppConfig;
    })(Configs = CognitiveServiceAddin.Configs || (CognitiveServiceAddin.Configs = {}));
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
        var TranslatorController = (function () {
            function TranslatorController($scope, $log, $timeout, officeService, translatorService) {
                this.$scope = $scope;
                this.$log = $log;
                this.$timeout = $timeout;
                this.officeService = officeService;
                this.translatorService = translatorService;
                var that = this;
            }
            TranslatorController.prototype.getSelectedText = function () {
                var that = this;
                that.officeService.getSelectedText()
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
            TranslatorController.prototype.translateSelectedText = function () {
                var that = this;
                that.translatorService.translateText(that.$scope.selectedText, "de")
                    .done(function (result) {
                    that.$timeout(0).then(function () {
                        that.$scope.translatedText = result;
                        console.log(result);
                    });
                })
                    .fail(function (error) {
                    console.error(error);
                });
            };
            TranslatorController.prototype.setText = function () {
                var that = this;
            };
            return TranslatorController;
        }());
        Controllers.TranslatorController = TranslatorController;
    })(Controllers = CognitiveServiceAddin.Controllers || (CognitiveServiceAddin.Controllers = {}));
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
//module InsertLocation {
//    var before: string;
//    var after: string;
//    var start: string;
//    var end: string;
//    var replace: string;
//} 
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Services;
    (function (Services) {
        var OfficeService = (function () {
            function OfficeService($log) {
                this.$log = $log;
                this.mailBoxItem = null;
                var that = this;
                that.mailBoxItem = Office.context.mailbox.item;
            }
            OfficeService.prototype.getSelectedText = function () {
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
    })(Services = CognitiveServiceAddin.Services || (CognitiveServiceAddin.Services = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Services;
    (function (Services) {
        var TranslatorService = (function () {
            function TranslatorService($log, $http) {
                this.$log = $log;
                this.$http = $http;
                var that = this;
            }
            TranslatorService.prototype.getAccessToken = function () {
                var that = this;
                var deferred = $.Deferred();
                that.$http({
                    method: 'POST',
                    url: CognitiveServiceAddin.Configs.AppConfig.AccessTokenUrl,
                    headers: {
                        'Ocp-Apim-Subscription-Key': CognitiveServiceAddin.Configs.AppConfig.TranslatorKey
                    }
                }).then(function (success) {
                    that.$log.debug("Got access token: ", success);
                    deferred.resolve(success.data);
                }, function (error) {
                    that.$log.error("Failed to get access token: ", error);
                    deferred.reject(error);
                });
                return deferred;
            };
            TranslatorService.prototype.translateText = function (text, to, from) {
                var that = this;
                var deferred = $.Deferred();
                var qureryString = '?text=' + text + '&to=' + to;
                if (from != null) {
                    qureryString = qureryString + '&from=' + from;
                }
                // get accessToken
                that.getAccessToken()
                    .done(function (response) {
                    // issue request to translate
                    that.$http({
                        method: 'GET',
                        url: 'https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20170616T132444Z.8a6358f25ffa04c9.4c7c24e4236de49256686cac785724cdc1e9f725&text=' + text + '&lang=en-ru&format=plain',
                        headers: {
                            'Content-Type': 'application/json',
                        }
                    }).then(function (success) {
                        that.$log.debug("Got access token: ", success);
                        deferred.resolve(success);
                    }, function (error) {
                        that.$log.error("Failed to get access token: ", error);
                        deferred.reject(error);
                    });
                })
                    .fail(function (error) {
                    deferred.reject(error);
                });
                return deferred;
            };
            return TranslatorService;
        }());
        Services.TranslatorService = TranslatorService;
    })(Services = CognitiveServiceAddin.Services || (CognitiveServiceAddin.Services = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
var CognitiveServiceAddin;
(function (CognitiveServiceAddin) {
    var Services;
    (function (Services) {
        var WordService = (function () {
            function WordService($log) {
                this.$log = $log;
                var that = this;
            }
            WordService.prototype.getSelectedText = function () {
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
            WordService.prototype.writeTextToMail = function (text) {
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
    })(Services = CognitiveServiceAddin.Services || (CognitiveServiceAddin.Services = {}));
})(CognitiveServiceAddin || (CognitiveServiceAddin = {}));
//# sourceMappingURL=app.js.map