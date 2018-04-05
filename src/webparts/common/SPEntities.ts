/**
 * Represents SharePoint List object
 */
export interface ISPList {
    Title: string;
    Id: string;
}

/**
 * Represents SharePoint REST service response for /_api/web/lists service call
 */
export interface ISPLists {
    value: ISPList[];
}

/**
 * Represents SharePoint View object
 */
export interface ISPView {
    Title: string;
    Id: string;
    ListId: string;
}

/**
 * Represents SharePoint REST service repsonse for /_api/web/lists(guid'<id>')/views service call
 */
export interface ISPViews {
    value: ISPView[];
}

export interface ISPField {
    Title: string;
    Id: string;
    ListId: string;
}

export interface ISPFields {
    value: ISPField[];
}