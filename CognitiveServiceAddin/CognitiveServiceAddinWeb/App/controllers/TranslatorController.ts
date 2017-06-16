module CognitiveServiceAddin.Controllers {
    interface ITranslatorCtrlScope  {
        selectedText?: string;
        translatedText?:string;
    }

    export class TranslatorController {
        constructor(private $scope: ITranslatorCtrlScope,
            private $log: ng.ILogService,
            private $timeout: ng.ITimeoutService,
            private officeService: Services.OfficeService,
            private translatorService: Services.TranslatorService) {
                var that: TranslatorController = this;
        }

        public getSelectedText() {
            var that: TranslatorController = this;

            that.officeService.getSelectedText()
                .done((result) => {
                    that.$timeout(0).then(() => {
                        that.$scope.selectedText = result;
                        console.log(result);                        
                    });
            })
            .fail((error) => {
                console.error(error);
            });
        }

        public translateSelectedText() {
            var that: TranslatorController = this;

            that.translatorService.translateText(that.$scope.selectedText, "de")
                .done((result) => {
                    that.$timeout(0).then(() => {
                        that.$scope.translatedText = result;
                        console.log(result);
                    });
                })
                .fail((error) => {
                    console.error(error);
                });
        }

        public setText() {
            var that: TranslatorController = this;

        }
    }
}