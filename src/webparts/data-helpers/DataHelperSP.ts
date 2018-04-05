import { ISPList, ISPView, ISPLists, ISPViews, ISPField, ISPFields } from "../common/SPEntities";
import { IDataHelper } from "./DataHelperBase";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

/**
 * List with views interface
 */
interface ISPListWithViews extends ISPList {
    /**
     * List Views
     */
    Views: ISPView[];
}

/**
 * SharePoint Data Helper class.
 * Gets information from current web
 */
export class DataHelperSP implements IDataHelper {
    /**
     * Web Part context
     */
    public context: IWebPartContext;
    /**
     * Loaded lists
     */
    private _lists: ISPListWithViews[];
    /**
     * constructor
     */
    public constructor(context: IWebPartContext) {
        this.context = context;
    }
    /**
     * API to get lists from the source
     */
    public getLists(): Promise<ISPList[]> {
        let queryUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`;
        return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            })
            .then((response: ISPLists) => {
                return response.value;
            });
    }
    /**
     * API to get views from the source
     */
    public getViews(listId: string): Promise<ISPView[]> {
        if (listId && listId == '-1' || listId == '0') {
            return new Promise<ISPView[]>((resolve) => {
                resolve(new Array<ISPView>());
            });
        }
    
        // try to get views from cache
        const lists: ISPListWithViews[] = this._lists && this._lists.length && this._lists.filter((value, index, array) => { 
            return value.Id === listId;
        });

        if (lists && lists.length) {
            return new Promise<ISPView[]>((resolve) => {
                resolve(lists[0].Views);
            });            
        }
        else {
            let queryUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/views`;
            return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    return response.json();
                })
                .then((response: ISPViews) => {
                    var views = response.value;
                    if (!this._lists || !this._lists.length){
                        this._lists = new Array<ISPListWithViews>();
                    }
                    this._lists.push({
                        Id: listId,
                        Title: '',
                        Views: views
                    });
                    return views;
                });
        }
    }
    /**
     * API to get fields from the source
     */
    public getFields(listId: string): Promise<ISPField[]> {
        let queryUrl = `${this.context.pageContext.web.absoluteUrl}/_api/list(guid'${listId}')/fields?$filter=Hidden eq false`;
        return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            })
            .then((response: ISPFields) => {
                return response.value;
            });
    }
}