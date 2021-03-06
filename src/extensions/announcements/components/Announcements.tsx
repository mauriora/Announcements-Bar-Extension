import * as React from 'react';
import { FunctionComponent, useCallback, useEffect, useMemo, useState } from 'react';
import { AnnouncementExtended } from '@mauriora/model-announcement-extended';
import { getCreateByIdOrTitle, ListItem, SharePointList, SharePointModel, getCurrentUser } from '@mauriora/controller-sharepoint-list';
import { ErrorBoundary, useAsyncError } from '@mauriora/utils-spfx-controls-react';
import { Spinner } from '@fluentui/react';
import { AnnouncementsList } from './AnnoncementsList';

export interface IAnnouncementsProps {
    siteUrl: string;
    listName: string;
    acknowledgedListName: string;
    culture: string;
}

export const ModelContext = React.createContext<SharePointModel<AnnouncementExtended>>(undefined);
ModelContext.displayName = 'ModelContext';

export const AcknkowledgedContext = React.createContext<SharePointModel<ListItem>>(undefined);
AcknkowledgedContext.displayName = 'AcknkowledgedContext';


const AnnouncementsLoader: FunctionComponent<IAnnouncementsProps> = ({ culture, listName, siteUrl, acknowledgedListName }) => {
    const [controller, setController] = useState<SharePointList>(undefined);
    const [announcements, setAnnouncements] = useState<SharePointModel<AnnouncementExtended>>(undefined);
    const [acknowledgements, setAcknowledgements] = useState<SharePointModel<ListItem>>(undefined);
    const throwError = useAsyncError();
    const currentUser = useMemo(() => getCurrentUser(''), []);

    console.log(`Announcements:AnnouncementsLoader render`, { currentUser, announcements, controller, acknowledgements, listName, siteUrl, culture });

    const getController = useCallback(
        async () => {
            try {
                const newController = await getCreateByIdOrTitle(listName, siteUrl);
                const now: string = new Date().toISOString();
                const newModel = await newController.addModel(
                    AnnouncementExtended,
                    `(StartDate le datetime'${now}' or StartDate eq null) and (Expires ge datetime'${now}' or Expires eq null)`
                );
                if(0 === newModel.records.length ) 
                {
                    await newModel.loadAllRecords();
                }

                setController(newController);
                setAnnouncements(newModel);
            } catch (controllerError) {
                throwError(controllerError);
            }
        },
        [listName, siteUrl]
    );

    const getAcknowledgements = useCallback(
        async () => {
            try {
                const newController = await getCreateByIdOrTitle(acknowledgedListName, siteUrl);

                const newModel = await newController.addModel(
                    ListItem,
                    `(Author/EMail eq '${currentUser.UserPrincipalName}')`
                );
                if(0 === newModel.records.length ) 
                {
                    await newModel.loadAllRecords();
                }
                setAcknowledgements(newModel);
            } catch (controllerError) {
                throwError(controllerError);
            }
        },
        [listName, siteUrl]
    );

    useEffect(() => { getController(); }, [listName, siteUrl]);
    useEffect(() => { getAcknowledgements(); }, [acknowledgedListName, siteUrl]);

    return announcements && acknowledgements ?
        <AcknkowledgedContext.Provider value={acknowledgements}>
            <ModelContext.Provider value={announcements}>
                <AnnouncementsList culture={culture} />
            </ModelContext.Provider>
        </AcknkowledgedContext.Provider>
        :
        <Spinner />;
};

/**
 * Create global ErrorBoundary with AnnouncementsLoader as child.
 * @param props are passed to AnnouncementsLoader
 * @returns 
 */
export const Announcements: FunctionComponent<IAnnouncementsProps> = props =>
    <ErrorBoundary>
        <AnnouncementsLoader {...props} />
    </ErrorBoundary>;