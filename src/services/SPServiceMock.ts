import { ISPService } from "./ISPService";
import { ISPWeb } from "../common/SPEntities";
import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";

export default class SPServiceMock implements ISPService {
  private readonly _context: IWebPartContext;

  constructor(context: IWebPartContext) {
    this._context = context;
  }

  public getRootWeb(): Promise<string> {
    return new Promise<string>(resolve => resolve(this._context.pageContext.web.absoluteUrl));
  }

  public getSites(rootWebUrl: string): Promise<ISPWeb[]> {
    const _sites: ISPWeb[] = [{Title: this._context.pageContext.web.title, Url: this._context.pageContext.web.absoluteUrl}];
    return new Promise<ISPWeb[]>(resolve => resolve(_sites));
  }
}
