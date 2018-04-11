import { ISPService } from "./ISPService";
import { ISPWeb } from "../common/SPEntities";
import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";

export default class SPServiceMock implements ISPService {
  public getRootWeb(): Promise<string> {
    return Promise.resolve('https://localhost/workbench');
  }

  public getSites(rootWebUrl: string): Promise<ISPWeb[]> {
    const _sites: ISPWeb[] = [{Title: 'Workbench', Url: 'https://localhost/workbench'}];
    return Promise.resolve(_sites);
  }
}
