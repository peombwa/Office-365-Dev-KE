module TimeFinder.Services {
    export class GraphService {
        constructor(private $log: ng.ILogService,
            private $http: ng.IHttpService,
            private authService: Services.AuthService) {
            var that: GraphService = this;
        }

        public findMeetingTimes(suggestionRequest: Models.IMeetingTimesRequest): JQueryDeferred<Models.IMeetingTimeSuggestionsResult> {
            var that: GraphService = this;
            var deferred = $.Deferred();

            // get cached token
            that.authService.getToken('aad')
                .done((token: string) => {
                    // make http post request
                    that.$http({
                            method: 'POST',
                            url: Configs.AppConfig.GraphEndPoint + '/me/findMeetingTimes',
                            data: JSON.stringify(suggestionRequest),
                            headers: {
                                'Content-Type': 'application/json',
                                'Authorization': 'Bearer ' + token
                            }
                        })
                        .then((successresponse: any) => {
                            deferred.resolve(successresponse.data);
                            },
                            (errorResponse) => {
                                deferred.reject(errorResponse);
                            });
                })
                .fail((error) => {
                    that.$log.error("Failed to fetch token");
                    deferred.reject(null);
                });
            return deferred;
        }

        public listUsers(): JQueryDeferred<Models.IUser[]> {
            var that: GraphService = this;
            var deferred = $.Deferred();
            that.authService.getToken('aad')
                .done((token: string) => {

                    that.$http({
                            method: 'GET',
                            url: Configs.AppConfig.GraphEndPoint+'/users/',
                            headers: {
                                'Content-Type': 'application/json',
                                'Authorization': 'Bearer ' + token
                            }
                        })
                        .then((successresponse: any) => {
                            that.$log.debug("successresponse: ", successresponse);
                                deferred.resolve(successresponse.data.value);
                            },
                            (errorResponse) => {
                                deferred.reject(errorResponse);
                            });
                });

            return deferred;
        }
    }
}