module CognitiveServiceAddin.Controllers {
    export class TranslatorController {
        constructor(private $scope: ng.IScope,
            private $log: ng.ILogService,
            private OfficeService: Services.OfficeService) {
                var that: TranslatorController = this;
        }
    }
}