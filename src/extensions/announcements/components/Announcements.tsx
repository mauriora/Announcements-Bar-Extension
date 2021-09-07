import * as React from 'react';
import { FunctionComponent, useCallback, useEffect, useMemo, useState } from 'react';
import { RateableAnnouncement } from '../models/RateableAnnouncement';
import { create as createController, ListItem, SharePointList, SharePointModel, getCurrentUser } from '@mauriora/controller-sharepoint-list';
import { ErrorBoundary, useAsyncError } from '@mauriora/utils-spfx-controls-react';
import { Spinner } from '@fluentui/react';
import { AnnouncementsList } from './AnnoncementsList';

export interface IAnnouncementsProps {
    siteUrl: string;
    listName: string;
    acknowledgedListName: string;
    culture: string;
}

export const ModelContext = React.createContext<SharePointModel<RateableAnnouncement>>(undefined);
ModelContext.displayName = 'ModelContext';

export const AcknkowledgedContext = React.createContext<SharePointModel<ListItem>>(undefined);
AcknkowledgedContext.displayName = 'AcknkowledgedContext';


const AnnouncementsLoader: FunctionComponent<IAnnouncementsProps> = ({ culture, listName, siteUrl, acknowledgedListName }) => {
    const [controller, setController] = useState<SharePointList>(undefined);
    const [model, setModel] = useState<SharePointModel<RateableAnnouncement>>(undefined);
    const [acknowledgements, setAcknowledgements] = useState<SharePointModel<ListItem>>(undefined);
    const throwError = useAsyncError();
    const currentUser = useMemo( () => getCurrentUser(''), [] );

    console.log(`Announcements:AnnouncementsLoader render`, { currentUser, controller, listName, siteUrl, culture });

    const getController = useCallback(
        async () => {
            try {
                const newController = listName && siteUrl ?
                    await createController(
                        listName, siteUrl
                    ) : 
                    undefined;
                
                console.log(`Announcements:AnnouncementsLoader.getController controller.init`, { newController, listName, siteUrl, currentUser, culture });
                await newController.init();

                const now: string = new Date().toISOString();
                const newModel = await newController.addModel(
                    RateableAnnouncement,
                    `(StartDate le datetime'${now}' or StartDate eq null) and (Expires ge datetime'${now}' or Expires eq null)`
                );
                console.log(`Announcements:AnnouncementsLoader.getController model.loadAllRecords()`, { newModel, newController, listName, siteUrl, currentUser, culture, now });
                await newModel.loadAllRecords();
                console.log(`Announcements:AnnouncementsLoader.getController setting context`, { newModel, newController, listName, siteUrl, currentUser, culture });

                setController(newController);
                setModel(newModel);
            } catch (controllerError) {
                throwError(controllerError);
            }
        },
        [listName, siteUrl]
    );

    const getAcknowledgements = useCallback(
        async () => {
            try {
                const newController = acknowledgedListName && siteUrl ?
                    await createController( acknowledgedListName, siteUrl ) : undefined;
                
                console.log(`Announcements:AnnouncementsLoader.getAcknowledgements controller.init`, { newController, acknowledgedListName, currentUser, siteUrl });
                await newController.init();

                const newModel = await newController.addModel(
                    ListItem,
                    `(Author/EMail eq '${currentUser.UserPrincipalName}')`
                );
                console.log(`Announcements:AnnouncementsLoader.getAcknowledgements model.loadAllRecords()`, { newModel, newController, acknowledgedListName, currentUser, siteUrl });
                await newModel.loadAllRecords();
                console.log(`Announcements:AnnouncementsLoader.getAcknowledgements setting context`, { newModel, newController, acknowledgedListName, currentUser, siteUrl });
                setAcknowledgements(newModel);
            } catch (controllerError) {
                throwError(controllerError);
            }
        },
        [listName, siteUrl]
    );

    useEffect(() => { getController(); }, [listName, siteUrl]);
    useEffect(() => { getAcknowledgements(); }, [acknowledgedListName, siteUrl]);

    return controller && model && acknowledgements ?
        <AcknkowledgedContext.Provider value={acknowledgements}>
            <ModelContext.Provider value={model}>
                <AnnouncementsList culture={culture} />
            </ModelContext.Provider>
        </AcknkowledgedContext.Provider>
        :
        <Spinner />;
};

export const Announcements: FunctionComponent<IAnnouncementsProps> = props => 
    <ErrorBoundary>
        <AnnouncementsLoader {...props} />
    </ErrorBoundary>;