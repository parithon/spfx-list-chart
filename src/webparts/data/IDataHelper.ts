import Guid from "@microsoft/sp-core-library/lib/Guid";
import { ISPField } from "../common/ISPEntities";

export interface IDataHelper {
  /**
   * API to get a SharePoint lists fields from the source
   */
  getFields(listId: Guid): Promise<ISPField[]>;
}
