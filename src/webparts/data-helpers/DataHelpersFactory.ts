import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IDataHelper } from "./DataHelperBase";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { DataHelperMock } from "./DataHelperMock";
import { DataHelperSP } from "./DataHelperSP";

export class DataHelpersFactory {
    /**
     * API to create data helper
     * @context: Web Part context
     */
    public static createDataHelper(context: IWebPartContext): IDataHelper {
        if (Environment.type === EnvironmentType.Local) {
            return new DataHelperMock();
        } else {
            return new DataHelperSP(context);
        }
    }
}