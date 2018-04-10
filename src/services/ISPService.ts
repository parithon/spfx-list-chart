import { ISPWeb } from "../common/SPEntities";

export interface ISPService {
  getRootWeb(): Promise<string>;
  getSites(rootWebUrl: string): Promise<ISPWeb[]>;
}
