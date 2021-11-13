# Source

This is derived from [Announcements SharePoint Framework Application Customizer](https://github.com/pnp/sp-dev-fx-extensions/tree/master/samples/react-application-announcements)

[Please refer to the root workspace documentation](../../README.md)

## Content

This submodule produces the Announcements-Bar-Extension for SharePoint.
The output can be found in [./sharepoint/solution/announcements-bar.sppkg](./sharepoint/solution/announcements-bar.sppkg)

### Code

All code is in [src/extensions/announcements](src/extensions/announcements).

- [AnnouncementsBar.ts](src/extensions/announcements/AnnouncementsBar.ts) extension entry point
- [components/Announcements.tsx](./src/extensions/announcements/components/Announcements.tsx) infrastructure and loading of Models
- [components/AnnouncementsList.tsx](./src/extensions/announcements/components/AnnouncementsList.tsx) render the announcements.

#### AnnouncementsBar entry

The entry point class component is 
__[AnnouncementsBar.ts](src/extensions/announcements/AnnouncementsBar.ts)__ . `AnnouncementsBar` is extended from `BaseApplicationCustomizer`.

The point of interest is `onInit()` creating [components/Announcements.tsx](./src/extensions/announcements/components/Announcements.tsx):

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

#### Announcements

[components/Announcements.tsx](./src/extensions/announcements/components/Announcements.tsx) exports the function component `Announcements`. It creates the global `ErrorBoundary` with `AnnouncementsLoader` as child.

`AnnouncementsLoader` shows a spinner until the models are loaded. Then it creates react contexts for acknkowledged announcements and the announcements. The child of the contexts is the `AnnouncementsList`.

#### AnnouncementsList

[components/AnnouncementsList.tsx](./src/extensions/announcements/components/AnnouncementsList.tsx) exports the function component `AnnouncementsList`. It handles the acknowlegment of Announcements and renders a stack of `AnnouncementContent`.
