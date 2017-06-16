module TimeFinder.Controllers {
    interface IUserVM extends Models.IUser {
        isSelected?:boolean;
    }
    interface IMainCtrlScope extends ng.IScope {
        adUsers?: IUserVM[];
        meetingTimeSuggestions?:any;
    }
    export class MainController {
        constructor(private $log: ng.ILogService,
            private $scope: IMainCtrlScope,
            private $timeout:ng.ITimeoutService,
            private graphService: Services.GraphService) {
            var that: MainController = this;      

        }

        public findMeetingTimes() {
            var that: MainController = this;
            var attendees: Models.IAttendee[] = [];

            angular.forEach(that.$scope.adUsers, (value, key) => {
                attendees.push({
                    type: "required",
                    emailAddress: {
                        address: value.mail,
                        name:value.displayName
                    }
                });
            });

            // set time constraiint

            var timeConstraint = {
                activityDomain: "unrestricted",
                timeslots:[
                    {
                        start: {
                            //2017-04-17T09:00:00
                            dateTime: moment().format(),
                            timeZone: "UTC"
                        },
                        end: {
                            dateTime: moment().add(5, "hours").format(),
                            timeZone: "UTC"
                        }
                    }]
            };

            that.graphService.findMeetingTimes({ attendees: attendees, timeConstraint: timeConstraint,  minimumAttendeePercentage: 10, returnSuggestionReasons: true})
                .done((response) => {
                    that.$log.debug("Response: ", response);
                    that.$timeout(0).then(() => {
                        that.$scope.meetingTimeSuggestions = response.meetingTimeSuggestions;
                    });                    
                })
                .fail((error) => {
                    that.$log.error("Error: ", error);
                });
        }

        public listADUsers() {
            var that: MainController = this;

            that.graphService.listUsers()
                .done((response) => {
                    that.$log.debug("Users: ", response);

                    that.$timeout(0).then(() => {
                        that.$scope.adUsers = response; 
                    });                    
                })
                .fail((error) => {
                    that.$log.error("Error: ", error);
                });
        }
    }
}