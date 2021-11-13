import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import * as Controller from '@mauriora/controller-sharepoint-list';
import { IAnnouncementsProps, Announcements } from './components/Announcements';
import { configure } from 'mobx';

export const QUALIFIED_NAME = 'Extension.ApplicationCustomizer.RateableAnnouncements';

export interface IAnnouncementsBarProperties {
    siteUrl: string;
    listName: string;
    acknowledgedListName: string;
}

/**
 * Mobx Configuration
 */
configure({
    enforceActions: "never"
});

export default class AnnouncementsBar extends BaseApplicationCustomizer<IAnnouncementsBarProperties> {

    protected async onInit(): Promise<void> {
        /** Print manifest information to console: alias, id, version */
        console.log(
            `${this?.context?.manifest?.alias} [${this?.context?.manifest?.id}] version=${this?.context?.manifest?.version} onInit start`,
            {
                properties: this.properties,
                context: this.context,
                me: this
            });

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
            console.error(
                `${this?.context?.manifest?.alias} [${this?.context?.manifest?.id}] version=${this?.context?.manifest?.version} onInit Controller.init caught` +
                `${(err as Error).message ?? err}` 
                );
        }

        /** find DOM element to render Announcement Bar on */
        const header = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

        if (!header) {
            const message = `${this?.context?.manifest?.alias} [${this?.context?.manifest?.id}] version=${this?.context?.manifest?.version} onInit Could not find placeholder Top`;
            console.error(message, { context: this.context, properties: this.properties });
            // return Promise.reject(new Error(message));
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

        console.log(`${this?.context?.manifest?.alias} [${this?.context?.manifest?.id}] version=${this?.context?.manifest?.version} onInit finished`, { propertiesDeconstructed: { ...this.properties }, properties: this.properties, context: this.context, contextDeconstructed: { ...this.context } });
    }
}
