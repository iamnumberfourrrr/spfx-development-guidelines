import { override } from "@microsoft/decorators";
import * as Rx from 'rx-lite';
import { ConfigurationListConstants } from "../CommonConstants";
import { log } from "../Utils";
import { BaseListService } from "./BaseListService";
import { NavService } from "./NavService";

export interface IConfigurationItem {
    Title: string;
    ConfigurationValue: string;
}

export interface IConfigurationPartialMatch {
    items: IConfigurationItem[];
}

export interface IConfigurationKeyCondition {
    keyPrefix: string;
}

export class ConfigurationService extends BaseListService {
    private _configurationItems: Rx.BehaviorSubject<IConfigurationItem[]> = new Rx.BehaviorSubject([]);
    private _loadedConfigurationItems = this._configurationItems.filter(cfgs => cfgs.length > 0);
    public Key = 'ConfigurationService';    

    @override
    protected environmentChecked() {
        this._loadConfiguration();
    }

    public async getConfiguration(...keys: (string | IConfigurationKeyCondition)[]): Promise<(IConfigurationItem | IConfigurationPartialMatch)[]> {
        return new Promise<(IConfigurationItem | IConfigurationPartialMatch)[]>((resolve) => {
            this._loadedConfigurationItems.take(1).subscribe(loadedCfgs => {
                const items = keys.map(key => this._getConfigurationItem(loadedCfgs, key));
                resolve(items);
            });
        });
    }

    private _getConfigurationItem(items: IConfigurationItem[], key: string | IConfigurationKeyCondition): IConfigurationItem | IConfigurationPartialMatch {
        const keyConfig = (key as IConfigurationKeyCondition);
        const partialMatch = Boolean(keyConfig.keyPrefix);
        var matchItem = items.filter((item) => {
            if (partialMatch) {
                return item.Title.indexOf(keyConfig.keyPrefix) >= 0;
            }
            return item.Title === key;
        });

        let result = null;
        if (partialMatch) {
            result = { items: matchItem } as IConfigurationPartialMatch;
        } else {
            result = matchItem.length > 0 ? matchItem[0] : null;
        }

        if (!result || result.length && result.length === 0) {
            console.warn('Missing configuration value for item ' + key);
        }
        return result;
    }

    private async _loadConfiguration() {
        log('Loading configurations');
        const configSite = NavService.getConfigSiteUrl();
        const result = await this.getListItems<IConfigurationItem>(configSite, ConfigurationListConstants.Configurations);
        this._configurationItems.onNext(result);
    }    
}