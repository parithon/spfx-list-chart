import { IDataHelper } from "./DataHelperBase";
import { ISPList, ISPView, ISPField } from "../common/SPEntities";

export class DataHelperMock implements IDataHelper {
    /**
     * hardcoded collection of lists
     */
    private static _lists: ISPList[] = [
        { Title: 'Test 1', Id: '1' },
        { Title: 'Test 2', Id: '2' },
        { Title: 'Test 3', Id: '3' }
    ];

    private static _views: ISPView[] = [
        { Title: 'All Items', Id: '1', ListId: '1' },
        { Title: 'Demo', Id: '2', ListId: '1' },
        { Title: 'All Items', Id: '1', ListId: '2' },
        { Title: 'All Items', Id: '1', ListId: '3' }
    ];

    private static _fields: ISPField[] = [
        { Title: 'Demo Column One', Id: '1', ListId: '1' },
        { Title: 'Demo Column Two', Id: '2', ListId: '1' },
        { Title: 'Demo Column Uno', Id: '1', ListId: '2' },
        { Title: 'Demo Column Un', Id: '1', ListId: '3' }
    ];

    /**
     * API to get lists from the source
     */
    public getLists(): Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(DataHelperMock._lists);
        });
    }

    /**
     * API to get views from the source
     */
    public getViews(listId: string): Promise<ISPView[]> {
        return new Promise<ISPView[]>((resolve) => {
            const result: ISPView[] = DataHelperMock._views.filter((value,index, array) => {
                return value.ListId === listId;
            });
            resolve(result);
        });
    }

    /**
     * API to get fields from the source
     */
    public getFields(listId: string): Promise<ISPField[]> {
        switch (listId) {
            case '6770c83b-29e8-494b-87b6-468a2066bcc6':
                listId = '1';
                break;
            case '2ece98f2-cc5e-48ff-8145-badf5009754c':
                listId = '2';
                break;
            case 'bd5dbd33-0e8d-4e12-b289-b276e5ef79c2':
                listId = '3';
                break;
            default:
                listId = '1';
                break;
        }
        return new Promise<ISPField[]>((resolve) => {
            this._sleep(500).then(() => {
                const result: ISPField[] = DataHelperMock._fields.filter((value, index, array) => {
                    return value.ListId === listId;
                });
                resolve(result);
            });
        });
    }

    private _sleep(milliseconds: number): Promise<void> {
        return new Promise((resolve) => setTimeout(resolve, milliseconds));
    }
}