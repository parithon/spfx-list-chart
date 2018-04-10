import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";
import { SPHttpClient } from "@microsoft/sp-http";

import { ISPService } from "./ISPService";
import { ISPWeb } from "../common/SPEntities";

export default class SPService implements ISPService {
  private readonly _context: IWebPartContext;

  constructor(context: IWebPartContext) {
    this._context = context;
  }

  public getRootWeb(): Promise<string> {
    let queryUrl: string = `${this._context.pageContext.web.absoluteUrl}/_api/site?$select=Url`;
    return this._context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then(response => response.json())
      .then(response => response.Url);
  }

  public getSites(rootWebUrl: string): Promise<ISPWeb[]> {
    const _sites: ISPWeb[] = [];
    let _queryUrl: string = `${rootWebUrl}/_api/web?select=Title,Url`;
    return this._context.spHttpClient.get(_queryUrl, SPHttpClient.configurations.v1)
      .then(rootWebResponse => rootWebResponse.json())
      .then((rootWeb: ISPWeb) => {
        _sites.push({
          Title: rootWeb.Title,
          Url: rootWeb.Url,
          IsRootWeb: true
        });
        _queryUrl = `${rootWebUrl}/_api/web/Webs?$select=Title,Url,effectivebasepermissions&$filter=effectivebasepermissions/high gt 32`;
        return this._context.spHttpClient.get(_queryUrl, SPHttpClient.configurations.v1)
          .then(response => response.json())
          .then(subSites => {
            subSites.value.forEach(subSite => {
              _sites.push({
                Title: subSite.Title,
                Url: subSite.Url
              });
            });
            return _sites;
          });
      });
  }
}
