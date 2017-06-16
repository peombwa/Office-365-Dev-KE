module CognitiveServiceAddin.Services {
    interface IOfficeBody extends Office.Body {
        getAsync?: any;
    }
    interface IOfficeItem extends Office.Item{
        body?: IOfficeBody;
        
    }

    export class OfficeService {
        private mailBoxItem: IOfficeItem = null;

        constructor(private $log: ng.ILogService) {
            var that: OfficeService = this;

            that.mailBoxItem = Office.context.mailbox.item;

        }

        public getSelectedText(): JQueryDeferred<string> {
            var that: OfficeService = this;
            var deferred = $.Deferred();


             that.mailBoxItem.body.getAsync("text",
                (result) => {
                    if (result.status == "succeeded") {
                        deferred.resolve(result.value);
                    } else {
                        deferred.reject(result);
                    }
                });

            return deferred;
        }

        public writeTextToMail(text: string): JQueryDeferred<boolean>  {
            var that: OfficeService = this;
            var deferred = $.Deferred();

            that.mailBoxItem.body.prependAsync(text,
                { coercionType: Office.CoercionType.Text, asyncContext: { var3: 1, var4: 2 } },
                (asyncResult) => {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        that.$log.error("Failed to add text to the body: ", asyncResult.error.message);
                        deferred.reject(asyncResult.error.message);
                    } else {
                        that.$log.debug("Added text to the body: ", that.mailBoxItem.body);
                        deferred.resolve(true);
                    }
                });

            return deferred;
        }
    
    }
}