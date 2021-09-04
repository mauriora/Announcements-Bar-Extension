import * as React from 'react';
import { FunctionComponent, useCallback, useContext, useEffect, useMemo, useState } from 'react';
import * as strings from 'announcementsStrings';
import { RateableAnnouncement } from '../models/RateableAnnouncement';
import { fromUserLookup, CommentsField, LikesCountField, PersonaHoverCard, RatingField, UserPersona } from '@mauriora/utils-spfx-controls-react';
import { Stack, MessageBar, MessageBarType, Link, Text, StackItem, PersonaSize, IMessageBarStyles, IStackTokens } from '@fluentui/react';
import { observer } from 'mobx-react-lite';
import { AcknkowledgedContext, ModelContext } from './Announcements';

interface RenderAnnouncementsProps {
    culture: string;
}

interface AnnouncementContentProps {
    announcement: RateableAnnouncement;
}

const messageBarStyles: IMessageBarStyles = {
    innerText: {
        flexGrow: 1
    },
    text: {
        flexGrow: 1
    }
};

const stackTokens: IStackTokens = {
    childrenGap: 's1',
    padding: 's1',
};

const AnnouncementContent: FunctionComponent<AnnouncementContentProps> = observer(({ announcement }) =>
    <Stack horizontal horizontalAlign='space-between'>
        {announcement.contentOwner &&
            <StackItem>
                <PersonaHoverCard user={fromUserLookup(announcement.contentOwner)} sendEmailButtonText={strings.SendEmailButton}>
                    <UserPersona
                        user={fromUserLookup(announcement.contentOwner)}
                        size={PersonaSize.size24}
                        imageUrl={announcement.contentOwner.picture}
                        imageAlt={announcement.contentOwner.title}
                        text={announcement.contentOwner.title}
                        showSecondaryText={false}
                    // secondaryText={announcement.contentOwner.claims ? announcement.contentOwner.claims.split('|').pop() : undefined}
                    />
                </PersonaHoverCard>
            </StackItem>
        }
        <StackItem>
            <Text variant='large'>{announcement.title}</Text>&nbsp;
        </StackItem>
        <StackItem >
            {/* 
                Unsafe set of HTML, this could cause XSS, use with care.
                Since the source list is under administrative control, this should be safe.
            */}
            <span style={{ whiteSpace: 'normal' }} dangerouslySetInnerHTML={{ __html: announcement.body }} />
        </StackItem>
        {announcement.url &&
            <StackItem>
                <Link href={announcement.url.url} target="_blank">{announcement.url.description ?? announcement.url.url}</Link>
            </StackItem>
        }
    </Stack>
);

export const AnnouncementsList: FunctionComponent<RenderAnnouncementsProps> = observer(({  }) => {    
    const model = useContext(ModelContext);
    const acknowledgedModel = useContext(AcknkowledgedContext);
    const [acknowledgedAnnouncements, setAcknowledgedAnnouncements] = useState<number[]>([]);
    const announcements = useMemo(
        () => model.records
            .filter(announcement => acknowledgedAnnouncements.indexOf(announcement.id) < 0),
        [model.records, model.records.length, acknowledgedAnnouncements, acknowledgedAnnouncements.length]
    );

    useEffect(() => {
        if(acknowledgedModel.records.length) {
            const items: number[] = JSON.parse(acknowledgedModel.records[0].title) || [];
            setAcknowledgedAnnouncements(items);
        }
    },[]);

    const onDismiss = useCallback(
        async (id: number) => {
            // Remove acknowleged announcement IDs that don't exist anymore, e.g. the announcement is expired or has been deleted
            const filteredAcknowledgedAnnouncements = acknowledgedAnnouncements.filter(
                prospect => model.records.some( announcement => announcement.id === prospect )
            );
            filteredAcknowledgedAnnouncements.push(id);
            const jsonString = JSON.stringify(filteredAcknowledgedAnnouncements);
            const record = acknowledgedModel.records.length ? acknowledgedModel.records[0] : acknowledgedModel.newRecord;
            record.title = jsonString;
            await acknowledgedModel.submit(record);

            setAcknowledgedAnnouncements(filteredAcknowledgedAnnouncements);
        },
        [acknowledgedAnnouncements, acknowledgedModel]
    );

    return <Stack tokens={stackTokens}>
        {announcements.map(announcement =>
            <StackItem>
                <MessageBar
                    messageBarType={(announcement.urgent ? MessageBarType.error : MessageBarType.warning)}
                    isMultiline={false}
                    onDismiss={() => onDismiss(announcement.id)}
                    dismissButtonAriaLabel={strings.Close}
                    styles={messageBarStyles}
                    actions={
                        <Stack horizontal>
                            <StackItem>
                                <CommentsField
                                    item={announcement}
                                    newCommentPlaceholder={strings.AddCommentPlaceholder}
                                    model={model}                                    
                                    property=''
                                    info={undefined}
                                    commentedText={strings.Commented}
                                />
                            </StackItem>
                            {'Likes' === announcement.controller.votingExperience ?
                                <StackItem>
                                    <LikesCountField mini model={model} item={announcement} info={model.propertyFields.get('likesCount')} property={'likesCount'} />
                                </StackItem>
                                :
                                'Ratings' === announcement.controller.votingExperience ?
                                    <StackItem>
                                        <RatingField model={model} item={announcement}  info={model.propertyFields.get('averageRating')} property={'averageRating'} />
                                    </StackItem>
                                    :
                                    undefined
                            }
                        </Stack>
                    }
                >
                    <AnnouncementContent announcement={announcement} />
                </MessageBar>
            </StackItem>
        )}
    </Stack>;
});
