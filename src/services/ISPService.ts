import { ISPWeb, ISPList } from "../common/SPEntities";
import ServiceScope from "@microsoft/sp-core-library/lib/serviceScope/ServiceScope";
import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";

export enum ListQueryOrderBy {
  Id = 1,
  Title
}
export interface ISPListQueryOptions {
  baseTemplate?: number;
  includeHidden?: boolean;
  orderBy?: ListQueryOrderBy;
}
export interface ISPService {
  getRootWeb(): Promise<string>;
  getSites(rootWebUrl: string): Promise<ISPWeb[]>;
  getLists(siteUrl: string, queryOptions?: ISPListQueryOptions): Promise<ISPList[]>;
}
