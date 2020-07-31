import { log } from "../Utils";
import { trimEnd } from "@microsoft/sp-lodash-subset";

let pageRenderedCallbacks: Function[] = [];
let pageNavigatedCallbacks: Function[] = [];

export enum Environment {
  DEV = "DEV",
  TEST = "TEST",
  PROD = "PROD",
  QA = "QA",
}

interface IEnvironmentConfig {
  portalSite: string;
  configSite: string;
  searchSite: string;
  webServiceUrl: string;
}
const environments: { [key: string]: IEnvironmentConfig } = {
  [Environment.DEV]: {
    portalSite: "/sites/projects",
    configSite: "/sites/config",
    searchSite: "/sites/search",
    webServiceUrl: "https://devoigsfcluster.theglobalfund.org",
  },
  [Environment.TEST]: {
    portalSite: "",
    configSite: "",
    searchSite: "",
    webServiceUrl: "",
  },
  [Environment.QA]: {
    portalSite: "",
    configSite: "",
    searchSite: "",
    webServiceUrl: "",
  },
  [Environment.PROD]: {
    portalSite: "",
    configSite: "",
    searchSite: "",
    webServiceUrl: "",
  },
};

export class NavService {
  public static getEnvironment() {
    // Logic to get env here
    // return the correct env
    log('Get environment', Environment.DEV);
    return Environment.DEV;
  }

  public static addPageRenderedCallback(cb: Function) {
    if (pageRenderedCallbacks.every((x) => x !== cb)) {
      pageRenderedCallbacks.push(cb);
    }
  }

  public static removePageRenderedCallback(cb: Function) {
    pageRenderedCallbacks = pageRenderedCallbacks.filter((x) => x !== cb);
  }

  public static notifyPageRenderdCallback() {
    pageRenderedCallbacks.forEach((c) => c());
  }

  public static addNavigatedCallback(cb: Function) {
    if (pageNavigatedCallbacks.every((x) => x !== cb)) {
      pageNavigatedCallbacks.push(cb);
    }
  }

  public static removeNavigatedCallback(cb: Function) {
    pageNavigatedCallbacks = pageNavigatedCallbacks.filter((x) => x !== cb);
  }

  public static notifyNavigatedCallback() {
    pageNavigatedCallbacks.forEach((c) => c());
  }

  public static getServerRelativeUrl(url: string) {
    if (
      !url ||
      (url.toLowerCase().indexOf("http:") < 0 &&
        url.toLowerCase().indexOf("https:") < 0)
    ) {
      return url || "";
    }
    const slashAfterHostHeader = url.indexOf("/", 8);
    return url.substr(slashAfterHostHeader);
  }

  public static getPathName(maxLevel?: number) {
    let pathName = trimEnd(location.pathname.toLocaleLowerCase(), "/");
    if (maxLevel) {
      return pathName
        .split("/")
        .slice(0, maxLevel + 1)
        .join("/");
    }
    return pathName;
  }

  private static isUrlMatch(serverRelativeUrl: string, pathName?: string) {
    pathName = pathName || NavService.getPathName(2);
    return serverRelativeUrl.toLocaleLowerCase().indexOf(pathName) >= 0;
  }

  public static parseParams(str: string) {
    let pieces = str.split("&"),
      data = {},
      i: number,
      parts: string[];
    for (i = 0; i < pieces.length; i++) {
      parts = pieces[i].split("=");
      if (parts.length < 2) {
        parts.push("");
      }
      const key = decodeURIComponent(parts[0]);
      const value = decodeURIComponent(parts[1]);
      if (value.indexOf(",") > -1) {
        data[key] = value.split(",");
      } else {
        data[key] = value;
      }
    }
    return data;
  }

  public static getSearchQueryUrlString() {
    return location.search.replace("?", "");
  }

  public static getHashUrlString() {
    return location.hash.replace("#", "");
  }

  public static getCurrentENVSearchPage() {
    return environments[this.getEnvironment()].searchSite;
  }

  public static getWebServiceBaseUrl() {
    return environments[this.getEnvironment()].webServiceUrl;
  }

  public static getPortalSiteUrl() {
    return environments[this.getEnvironment()].portalSite;
  }

  public static getConfigSiteUrl() {
    return environments[this.getEnvironment()].configSite;
  }
}
