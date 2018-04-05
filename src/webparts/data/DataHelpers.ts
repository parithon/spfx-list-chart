import { IDataHelper } from "./IDataHelper";
import Guid from "@microsoft/sp-core-library/lib/Guid";
import { ISPList, ISPLists, ISPField, ISPFields } from "../common/ISPEntities";
import { IWebPartContext } from "@microsoft/sp-webpart-base/lib";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http/lib";

export class MockDataHelper implements IDataHelper {
  private static _lists: ISPList[] = [
    {
      Id: Guid.parse('6770c83b-29e8-494b-87b6-468a2066bcc6'), // https://github.com/SharePoint/sp-dev-fx-property-controls/blob/master/src/services/SPListPickerService.ts
      Title: 'Mock List One'
    },
    {
      Id: Guid.parse('2ece98f2-cc5e-48ff-8145-badf5009754c'), // https://github.com/SharePoint/sp-dev-fx-property-controls/blob/master/src/services/SPListPickerService.ts
      Title: 'Mock List Two'
    },
    {
      Id: Guid.parse('bd5dbd33-0e8d-4e12-b289-b276e5ef79c2'), // https://github.com/SharePoint/sp-dev-fx-property-controls/blob/master/src/services/SPListPickerService.ts
      Title: 'Mock List Three'
    }
  ];
  private static _fields: ISPField[] = [
    {
      Id: Guid.parse('ae7b3abf-6010-4cdd-97c0-26ef5ec07433'),
      ListId: MockDataHelper._lists[0].Id,
      Title: 'Id'
    },
    {
      Id: Guid.parse('ae06f926-2e05-4481-a221-e4392a7a1387'),
      ListId: MockDataHelper._lists[0].Id,
      Title: 'Title'
    },
    {
      Id: Guid.parse('48d3bfc2-8570-4a41-9bb1-43002018b647'),
      ListId: MockDataHelper._lists[0].Id,
      Title: 'Mock List One Data'
    },
    {
      Id: Guid.parse('ab2abe7c-5ff6-49c0-8020-055a2f83b7e0'),
      ListId: MockDataHelper._lists[1].Id,
      Title: 'Id'
    },
    {
      Id: Guid.parse('6f6bb4ab-1427-418b-b3eb-dad2880ab179'),
      ListId: MockDataHelper._lists[1].Id,
      Title: 'Title'
    },
    {
      Id: Guid.parse('d617e917-92ba-4d40-a5c4-61d8db8748cc'),
      ListId: MockDataHelper._lists[1].Id,
      Title: 'Mock List Two Data'
    },
    {
      Id: Guid.parse('447f182f-a872-4a17-bd5f-16234a7c865c'),
      ListId: MockDataHelper._lists[2].Id,
      Title: 'Id'
    },
    {
      Id: Guid.parse('d3d75933-1803-4a68-b864-4c53a652e764'),
      ListId: MockDataHelper._lists[2].Id,
      Title: 'Title'
    },
    {
      Id: Guid.parse('0171cea2-d192-411b-8cdc-60411ff1f0b5'),
      ListId: MockDataHelper._lists[2].Id,
      Title: 'Mock List Three Data'
    }
  ];

  public getFields(listId: Guid): Promise<ISPField[]> {
    return new Promise<ISPField[]>((resolve) => {
      resolve(MockDataHelper._fields);
    });
  }
}

export interface ISPListWithFields extends ISPList {
  Fields: ISPField[];
}

export class SPDataHelper implements IDataHelper {
  private _lists: ISPListWithFields[] = [];
  constructor(context: IWebPartContext){
    this.context = context;
  }
  public context: IWebPartContext;
  public getFields(listId: Guid): Promise<ISPField[]> {
    if (!listId) {
      return new Promise<ISPField[]>((resolve) => {
        resolve(new Array<ISPField>());
      });
    }

    // Retrieve results from cache;
    const lists: ISPListWithFields[] = this._lists && this._lists.length && this._lists.filter((value, index, array) => {
      return value.Id === listId;
    });

    // If fields exist from cache, return the cached fields
    if (lists && lists.length) {
      return new Promise<ISPField[]>((resolve) => {
        resolve(lists[0].Fields);
      });
    }

    // Query for the fields
    let queryUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false`;
    return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((response: ISPFields) => {
        const fields = response.value;
        if (!this._lists || !this._lists.length) {
          this._lists = new Array<ISPListWithFields>();
        }
        this._lists.push({
          Id: listId,
          Title: '',
          Fields: fields
        });
        return fields;
      });
  }
}
