import { Announcement, Link, UserLookup } from '@mauriora/controller-sharepoint-list';
import { Expose, Type } from 'class-transformer';


export class RateableAnnouncement extends Announcement {
    @Expose({ name: 'Urgent' })
    public urgent: boolean;

    @Expose({ name: 'StartDate' })
    public startDate: string;

    @Type( () => Link )
    @Expose({name: 'URL'})
    public url: Link;

    @Type( () => UserLookup )
    @Expose({name: 'ReportOwner'})
    public contentOwner: UserLookup;
}