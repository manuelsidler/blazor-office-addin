window.wordWrapper = {
    getDocumentMetadata: async function () {
        const metadata = await window.Word.run(async context => {
            let properties = context.document.properties;
            context.load(properties);

            await context.sync();

            return {
                title: properties.title,
                subject: properties.subject
            };
        });

        return metadata;
    },
    saveDocumentMetadata: async function (metadata) {
        await window.Word.run(async context => {
            context.document.properties.title = metadata.title;
            context.document.properties.subject = metadata.subject;

            await context.sync();
        });
    }
}
