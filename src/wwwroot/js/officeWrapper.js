var wordWrapper = wordWrapper || {};

wordWrapper.getDocumentMetadata = async function () {
    try {
        const metadata = await Word.run(async context => {
            let properties = context.document.properties;
            context.load(properties);

            await context.sync();

            return {
                title: properties.title,
                subject: properties.subject
            };
        });

        return metadata;
    } catch (error) {
        throw Error(JSON.stringify(error));
    }
};

wordWrapper.saveDocumentMetadata = async function (metadata) {
    try {
        await Word.run(async context => {
            context.document.properties.title = metadata.title;
            context.document.properties.subject = metadata.subject;

            await context.sync();
        });
    } catch (error) {
        throw Error(JSON.stringify(error));
    }
};
