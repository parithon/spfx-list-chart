import { IWebPartContext } from "@microsoft/sp-webpart-base/lib";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library/lib";
import { IDataHelper } from "./IDataHelper";
import { MockDataHelper, SPDataHelper } from "./DataHelpers";

export class DataFactory {
  /**
   * API to create a data helper
   */
  public static createDataHelper(context: IWebPartContext): IDataHelper {
    if (Environment.type === EnvironmentType.Local) {
      return new MockDataHelper();
    }

    return new SPDataHelper(context);
  }
}
