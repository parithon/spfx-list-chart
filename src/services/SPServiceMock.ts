import { ISPService, ISPListQueryOptions } from "./ISPService";
import { ISPWeb, ISPList } from "../common/SPEntities";
import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";

const WORKBENCH_URL = 'https://wwww.contoso.com/sites/workbench';
export default class SPServiceMock implements ISPService {
  private static Sites: ISPWeb[] = [{
    Title: 'Workbench',
    Url: WORKBENCH_URL,
    Lists: [
      { Id: 'a9d0e259-11a9-45fc-979f-ad4af329a151', Title: 'Mock List One', BaseTemplate: 109 },
      { Id: '1969cfff-eefa-468a-b8de-9b88aa1dfcc6', Title: 'Mock List Two', BaseTemplate: 109 },
      { Id: '81c92fc0-dece-4b3e-9dda-21f1d707ac73', Title: 'Mock List Three', BaseTemplate: 109 }
    ]
  }];

  public getRootWeb(): Promise<string> {
    return Promise.resolve(WORKBENCH_URL);
  }

  public getSites(rootWebUrl: string): Promise<ISPWeb[]> {
    return Promise.resolve(SPServiceMock.Sites);
  }

  public getLists(siteUrl: string, queryOptions?: ISPListQueryOptions): Promise<ISPList[]> {
    const { Lists } = SPServiceMock.Sites[0];
    return Promise.resolve(Lists);
  }
}
