import { find } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { NavService } from './services/NavService';

export function getParentElement(dom: HTMLElement, match: (dom: HTMLElement) => boolean) {
    let parent = dom && dom.parentElement;
    while (parent) {
        if (match(parent)) {
            return parent;
        }
        parent = parent.parentElement;
    }
}

export function getUserImage(email: string): Promise<string> {
    return new Promise(resolve => {
        const url = `${NavService.getPortalSiteUrl()}/_layouts/15/userphoto.aspx?size=M&accountname=${email}`;
        const img = document.createElement('img') as HTMLImageElement;
        img.src = url;
        img.style.position = 'absolute';
        img.style.left = '-10000px';
        img.onload = () => {
            if (img.width === img.height && img.width < 5) {
                document.body.removeChild(img);
                resolve('');
            } else {
                document.body.removeChild(img);
                resolve(url);
            }
        };
        document.body.appendChild(img);
    });
}

export function getParent(dom: HTMLElement, cssClass: string) {
    return getParentElement(dom, d => d.classList.contains(cssClass));
}

export function getReactFromDOM(dom: HTMLElement) {
    let key = find(Object.keys(dom), k => k.indexOf('__reactInternalInstance$') === 0);
    let internalInstance = dom[key];
    if (!internalInstance) {
        return null;
    }

    return internalInstance.return ? internalInstance.return.stateNode : internalInstance._currentElement._owner._instance;
}

export function log(msg: string, ...args: any[]) {
    if (args && args.length > 0) {
        console.log(`ELCAVN - ${msg}`, args);
    } else {
        console.log(`ELCAVN - ${msg}`);
    }
}

function makeGlobalCacheKey(key) {
    return `ELCAVN_${key}`;
}

export function getGlobalCache<T = any>(key: string): { [key: string]: T } {
    const uniqkey = makeGlobalCacheKey(key);
    if (typeof (window[uniqkey]) === 'undefined') {
        window[uniqkey] = {};
    }
    return window[uniqkey];
}

export function setGlobalCache(key: string, val: any) {
    const uniqkey = makeGlobalCacheKey(key);
    if (!isGlobalCacheAvailable(key)) {
        window[uniqkey] = val;
    }
    window[uniqkey] = val;
}

export function isGlobalCacheAvailable(key: string) {
    const uniqkey = makeGlobalCacheKey(key);
    return !!window[uniqkey];
}