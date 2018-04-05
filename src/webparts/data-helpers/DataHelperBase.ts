import { ISPList, ISPView, ISPField } from '../common/SPEntities';

/**
 * Data Helpers interface
 */
export interface IDataHelper {
    /**
     * API to get lists from the source
     */
    getLists(): Promise<ISPList[]>;
    /**
     * API to get views from the source
     */
    getViews(listId: string): Promise<ISPView[]>;
    /**
     * API to get fields from the source
     */
    getFields(listId: string): Promise<ISPField[]>;
}