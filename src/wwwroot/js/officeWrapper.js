window.wordWrapper = {
    getDocumentMetadata: function () {
        return window.Word.run(function (context) {
            var properties = context.document.properties;
            context.load(properties);

            return context.sync()
                .then(function () {
                    return {
                        title: properties.title,
                        subject: properties.subject
                    };
                })
                .catch(function (error) {
                    return error;
                });
        });
    },
    saveDocumentMetadata: function(metadata) {
        return window.Word.run(function(context) {
            context.document.properties.title = metadata.title;
            context.document.properties.subject = metadata.subject;

            return context.sync();
        });
    }
}
