module YandexAddin.Services {
    export class WordService {
        constructor(private $log: ng.ILogService) {
            var that: WordService = this;

        }

        public getBodyText(): JQueryDeferred<string> {
            var that: WordService = this;
            var deferred = $.Deferred();

            Word.run((ctx) => {
                // create proxy object for the document
                var docBody = ctx.document.body;

                // queue command to load the text property of proxy object
                ctx.load(docBody, "text");
                // sync the doument object with the proxy objects
                return ctx.sync().then(() => {
                    that.$log.debug("Body contents: ", docBody.text);
                    deferred.resolve(docBody.text);
                }).catch((error) => {
                    that.$log.error("Failed to get text from the body: ", error);
                    deferred.reject(error);
                });

            });

            return deferred;
        }

        public writeTextToBody(text: string): JQueryDeferred<boolean> {
            var that: WordService = this;
            var deferred = $.Deferred();

            Word.run((ctx) => {
                var docBody = ctx.document.body;

                ctx.load(docBody, "text");

                docBody.insertText(text, Word.InsertLocation.end);

                return ctx.sync().then(() => {
                    that.$log.debug("New body contents: ", docBody.text);
                    deferred.resolve(true);
                }).catch((error) => {
                    that.$log.error("Failed to add text to the body: ", error);
                    deferred.reject(error);
                });
            });

            return deferred;
        }
    
    }
}