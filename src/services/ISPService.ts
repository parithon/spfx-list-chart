import { ISPWeb } from "../common/SPEntities";
import ServiceScope from "@microsoft/sp-core-library/lib/serviceScope/ServiceScope";
import IWebPartContext from "@microsoft/sp-webpart-base/lib/core/IWebPartContext";

export interface ISPService {
  getRootWeb(): Promise<string>;
  getSites(rootWebUrl: string): Promise<ISPWeb[]>;
}
