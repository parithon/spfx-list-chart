import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";
import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";

import { ISPService, ISPListQueryOptions, ListQueryOrderBy } from "./ISPService";
import { ISPWeb, ISPList } from "../common/SPEntities";

export default class SPService implements ISPService {
  public static readonly serviceKey: ServiceKey<SPService> = ServiceKey.create('gfins:SPService', SPService);

  private _serviceScope: ServiceScope;
  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _rootWebUrl: string;
  private _sites: ISPWeb[] = [];

  constructor(serviceScope: ServiceScope) {
    this._serviceScope = serviceScope;
    // this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
    // this.pageContext = serviceScope.consume(PageContext.serviceKey);
  }

  public getRootWeb(): Promise<string> {
    if (this._rootWebUrl && this._rootWebUrl.length > 0) {
      return Promise.resolve(this._rootWebUrl);
    }

    this.consumeDependencies();

    let queryUrl: string = `${this._pageContext.web.absoluteUrl}/_api/site?$select=Url`;
    return this._spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then(response => response.json())
      .then(response => {
        this._rootWebUrl = response.Url;
        return this._rootWebUrl;
      });
  }

  public getSites(rootWebUrl: string): Promise<ISPWeb[]> {
    if (this._sites.length > 0) {
      return Promise.resolve(this._sites);
    }

    this.consumeDependencies();

    let _queryUrl: string = `${rootWebUrl}/_api/web?$select=Title,Url`;
    return this._spHttpClient.get(_queryUrl, SPHttpClient.configurations.v1)
      .then(rootWebResponse => rootWebResponse.json())
      .then((rootWeb: ISPWeb) => {
        this._sites.push({
          Title: rootWeb.Title,
          Url: rootWeb.Url,
          IsRootWeb: true
        });
        _queryUrl = `${rootWebUrl}/_api/web/Webs?$select=Title,Url,effectivebasepermissions&$filter=effectivebasepermissions/high gt 32`;
        return this._spHttpClient.get(_queryUrl, SPHttpClient.configurations.v1)
          .then(response => response.json())
          .then(subSites => {
            subSites.value.forEach(subSite => {
              this._sites.push({
                Title: subSite.Title,
                Url: subSite.Url
              });
            });
            return this._sites;
          });
      });
  }

  public getLists(siteUrl: string, queryOptions?: ISPListQueryOptions): Promise<ISPList[]> {
    let site = this._sites.filter(site => site.Url === siteUrl)[0];
    const { Lists } = site;
    if (site !== null && Lists !== undefined) {
      return Promise.resolve(Lists);
    }

    this.consumeDependencies();

    let queryUrl: string = `${siteUrl}/_api/web/lists??$select=Title,id,BaseTemplate`;
    let filtered: boolean = false;

    if (queryOptions.baseTemplate !== undefined) {
      queryUrl += `&$filter=BaseTemplate eq ${queryOptions.baseTemplate}`;
      filtered = true;
    }

    if (queryOptions.includeHidden === false) {
      queryUrl += (filtered ? ' and Hidden eq false' : '&$filter=Hidden eq false');
      filtered = true;
    }

    if (queryOptions.orderBy !== undefined) {
      queryUrl += `&$orderby=${(queryOptions.orderBy === ListQueryOrderBy.Id ? 'Id': 'Title')}`
    }

    return this._spHttpClient.get(queryUrl, SPHttpClient.configurations.v1)
      .then(response => response.json())
      .then(listCollection => {
        if (site === null) {
          const idx = this._sites.push({ Url: siteUrl, Title: '' });
          site = this._sites[idx];
        }
        site.Lists = [];
        listCollection.map(list => {
          site.Lists.push({ Id: list.Id, Title: list.Title, BaseTemplate: list.BaseTemplate });
        })
        return site.Lists;
      });
  }

  private consumeDependencies() {
    if (this._spHttpClient === undefined) {
      this._spHttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);
    }
    if (this._pageContext === undefined) {
      this._pageContext = this._serviceScope.consume(PageContext.serviceKey);
    }
  }
}
