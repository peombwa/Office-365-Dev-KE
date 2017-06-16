module CognitiveServiceAddin.Services {
    export class TranslatorService {
        constructor(private $log: ng.ILogService,
            private $http: ng.IHttpService) {
            var that: TranslatorService = this;

        }

        private getAccessToken(): JQueryDeferred<string> {
            var that: TranslatorService = this;
            var deferred = $.Deferred();

            that.$http({
                method: 'POST',
                url: Configs.AppConfig.AccessTokenUrl,
                headers: {
                    'Ocp-Apim-Subscription-Key': Configs.AppConfig.TranslatorKey
                }
            }).then((success) => {
                    that.$log.debug("Got access token: ", success);
                    deferred.resolve(success.data);
                }, (error) => {
                    that.$log.error("Failed to get access token: ", error);
                    deferred.reject(error);
            });
            return deferred;
        }

        public translateText(text: string, to: string, from?: string): JQueryDeferred<string> {
            var that: TranslatorService = this;
            var deferred = $.Deferred();

            var qureryString: string = '?text='+ text + '&to=' + to;

            if (from != null) {
                qureryString = qureryString + '&from=' + from;
            } 

            // get accessToken
            that.getAccessToken()
                .done((response) => {
                    // issue request to translate
                    that.$http({
                        method: 'GET',
                        url: 'https://translate.yandex.net/api/v1.5/tr.json/translate?key=trnsl.1.1.20170616T132444Z.8a6358f25ffa04c9.4c7c24e4236de49256686cac785724cdc1e9f725&text='+text+'&lang=en-ru&format=plain',//Configs.AppConfig.TranslatorEndpoint + qureryString,
                        headers: {
                            'Content-Type': 'application/json',
                            //'Authorization': 'Bearer ' + response
                        }
                    }).then((success) => {
                        that.$log.debug("Got access token: ", success);
                        deferred.resolve(success);
                    }, (error) => {
                        that.$log.error("Failed to get access token: ", error);
                        deferred.reject(error);
                    });
                })
                .fail((error) => {
                    deferred.reject(error);
                });
          

            return deferred;
        }
    }
}