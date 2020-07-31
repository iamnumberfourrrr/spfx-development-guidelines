import { BaseComponentContext } from '@microsoft/sp-component-base';
import { getGlobalCache, log } from '../Utils';
import { BaseListService } from './BaseListService';
import { NavService } from './NavService';

export interface ServiceCtor<T> {
    new(context?: BaseComponentContext, serviceFactoryCreate?:boolean): T;
}

export class ServiceFactory {
    private static get serviceCache(): {
        [key: string]: { [key: string]: any };
    } {
        const key = 'ServiceFactory';
        return getGlobalCache(key);
    }

    public static getService<T extends BaseListService>(ctor: ServiceCtor<T>, context: BaseComponentContext, path?: string): T {
        const customUrl = !!path;
        const pathName = path || NavService.getPathName(2);
        const serviceInstance = new ctor(context, false);
        const key = serviceInstance.Key;
        log(`Service ${key} requested`);
        if (!this.serviceCache[pathName]) {
            this.serviceCache[pathName] = {};
        }

        if (!this.serviceCache[pathName][key]) {
            log('init new instance of service ', key);
            this.serviceCache[pathName][key] = serviceInstance;
            serviceInstance.init();
        }
        const cache = this.serviceCache[pathName][key] as T;
        if (cache.context !== context) {
            // refresh context (e.g after navigated)
            cache.context = context;
            cache.init();
        }
        if (customUrl) {
            cache.currentWebUrl = path;
        }
        return cache;
    }
}
