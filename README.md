# Source

This is derived from [Announcements SharePoint Framework Application Customizer](https://github.com/pnp/sp-dev-fx-extensions/tree/master/samples/react-application-announcements)

[Please refer to the root workspace documentation](../../README.md)

## Content

This submodule produces the Announcements-Bar-Extension for SharePoint.
The output can be found in [./sharepoint/solution/announcements-bar.sppkg](./sharepoint/solution/announcements-bar.sppkg)

### Code

__`[src/extensions/announcements/AnnouncementsBar.ts](src/extensions/announcements/AnnouncementsBar.ts)`__
Entry point Class Componenet extended from `BaseApplicationCustomizer`.
Point of interest is `onInit()`:

```typescript
    /** Print manifest information to console: alias, id, version */
    console.log( ... );

    /** Initialise Base class */
    super.onInit();

    /** Check we got all properties */
    if (!this.properties.siteUrl || !this.properties.listName || !this.properties.acknowledgedListName) {
        const message = `${this?.context?.manifest?.alias} [${this?.context?.manifest?.id}] version=${this?.context?.manifest?.version} onInit Missing required configuration parameters`;
        console.error(message, { context: this.context, properties: this.properties });
        // return Promise.reject(new Error(message));
    }

    /** Initialise SharePoint controller module with context */
    try {
        await Controller.init(this.context);
    } catch (err) {
        ...
    }

    /** find DOM element to render Announcement Bar on */
    const header = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

    if (!header) {
        ...
    }

    const site = this.context.pageContext.site;
    const tenantUrl = site.absoluteUrl.replace(site.serverRelativeUrl, "");

    /** Create Announcements React Element */
    const elem: React.ReactElement<IAnnouncementsProps> = React.createElement(Announcements, {
        siteUrl: `${tenantUrl}${this.properties.siteUrl}`,
        listName: this.properties.listName,
        acknowledgedListName: this.properties.acknowledgedListName,
        culture: this.context.pageContext.cultureInfo.currentUICultureName
    });

    /** Render Announcements on header dom element */
    ReactDOM.render(elem, header.domElement);
```
