module YandexAddin.Services {
    export class TranslatorService {
        constructor(private $log: ng.ILogService,
            private $http: ng.IHttpService) {
            var that: TranslatorService = this;

        }

        public translateText(text: string, to: string, from: string): JQueryDeferred<string> {
            var that: TranslatorService = this;
            var deferred = $.Deferred();

            var qureryString: string = '?key=' + Configs.AppConfig.TranslatorKey
                + '&text=' + text + '&lang=' + from + '-' + to + '&format=plain';

                // issue request to translate
                that.$http({
                    method: 'GET',
                    url: Configs.AppConfig.TranslatorEndpoint + qureryString,
                    headers: {
                        'Content-Type': 'application/json'
                    }
                }).then((success: any) => {
                    that.$log.debug("Successfully translated text: ", success);
                    deferred.resolve(success.data.text[0]);
                }, (error) => {
                    that.$log.error("Failed to translate text: ", error);
                    deferred.reject(error);
                });
                

            return deferred;
        }
    }
}