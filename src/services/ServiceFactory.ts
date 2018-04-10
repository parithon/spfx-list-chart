import IWebPartContext from '@microsoft/sp-webpart-base/lib/core/IWebPartContext';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

import { ISPService } from './ISPService';
import SPService from './SPService';
import SPServiceMock from '../../lib/services/SPServiceMock';

export default class ServiceFactory {
  public static createService(context: IWebPartContext): ISPService {
    if (Environment.type == EnvironmentType.Local) {
      return new SPServiceMock(context);
    }
    return new SPService(context);
  }
}
