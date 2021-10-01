import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import * as Controller from '@mauriora/controller-sharepoint-list';

import { IAnnouncementsProps, Announcements } from './components/Announcements';
import { configure } from 'mobx';

export const QUALIFIED_NAME = 'Extension.ApplicationCustomizer.RateableAnnouncements';

export interface IRateableAnnouncementsExtensionProperties {
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

export default class RateableAnnouncementsExtension
    extends BaseApplicationCustomizer<IRateableAnnouncementsExtensionProperties> {

    protected async onInit(): Promise<void> {
        console.log(`${this.context.manifest.alias} [${this.context.manifest.id}] version=${this.context.manifest.version} onInit super.onInit...`, { context: this.context, properties: this.properties });
        super.onInit();
        await Controller.init(this.context);

        if (!this.properties.siteUrl || !this.properties.listName || !this.properties.acknowledgedListName) {
            console.error(`${this.context.manifest.alias} [${this.context.manifest.id}] version=${this.context.manifest.version} onInit Missing required configuration parameters`, { context: this.context, properties: this.properties });
            const e: Error = new Error('Missing required configuration parameters');
            Log.error(QUALIFIED_NAME, e);
            return Promise.reject(e);
        }
        const header = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

        if (!header) {
            console.error(`${this.context.manifest.alias} [${this.context.manifest.id}] version=${this.context.manifest.version} onInit Could not find placeholder Top`, { context: this.context, properties: this.properties });
            const error = new Error('Could not find placeholder Top');
            Log.error(QUALIFIED_NAME, error);
            return Promise.reject(error);
        }

        const site = this.context.pageContext.site;
        const tenantUrl = site.absoluteUrl.replace(site.serverRelativeUrl, "");

        const elem: React.ReactElement<IAnnouncementsProps> = React.createElement(Announcements, { 
            siteUrl: `${tenantUrl}${this.properties.siteUrl}`, 
            listName: this.properties.listName,
            acknowledgedListName: this.properties.acknowledgedListName,
            culture: this.context.pageContext.cultureInfo.currentUICultureName
         });

        ReactDOM.render(elem, header.domElement);

        console.log(`${this.context.manifest.alias} [${this.context.manifest.id}] version=${this.context.manifest.version} onInit finished`, {propertiesDeconstructed: {...this.properties}, properties: this.properties, context: this.context, contextDeconstructed: {...this.context}});
    }
}
