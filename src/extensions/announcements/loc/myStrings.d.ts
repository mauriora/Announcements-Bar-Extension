declare interface IAnnouncementsStrings {
    Title: string;
    Close: string;
    Commented: string;
    AddCommentPlaceholder: string;
    SendEmailButton: string;
}

declare module 'announcementsStrings' {
    const strings: IAnnouncementsStrings;
    export = strings;
}
