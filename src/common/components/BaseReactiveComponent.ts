import { BaseComponentContext } from "@microsoft/sp-component-base";
import * as React from "react";
import * as Rx from 'rx-lite';

export interface BaseReactiveComponentProps {
    context: BaseComponentContext;
}

export abstract class BaseReactiveComponent<P extends BaseReactiveComponentProps = BaseReactiveComponentProps, S = {}> extends React.Component<P, S> {
    protected subs: Rx.IDisposable[] = [];
    protected done$ = new Rx.Subject();

    public componentWillUnmount() {
        this.componentDestroy();       
    }

    public componentDestroy() {
        this.subs.forEach(s => s.dispose());
        this.done$.onNext(undefined);
    }
}