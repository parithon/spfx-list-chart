import { Environment, EnvironmentType, ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import WebPartContext from '@microsoft/sp-webpart-base/lib/core/WebPartContext';

import { ISPService } from './ISPService';
import SPServiceMock from './SPServiceMock';
import SPService from './SPService';

const SERVICE_KEY_NAME: string = 'ListCharts.Services.SPServiceBase';

export default class ServiceFactory {
  public static createService(context: WebPartContext): ISPService {
    if (Environment.type === EnvironmentType.Local) {
      return new SPServiceMock();
    }
    const serviceScope: ServiceScope = context.serviceScope.getParent();
    return serviceScope.getParent().consume(SPService.serviceKey);
  }
}
