import { NavService } from './../../common/services/NavService';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from "@microsoft/sp-application-base";
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import * as strings from "ProjectInformationApplicationCustomizerStrings";
import ProjectInformation from "./components/ProjectInformation";
import styles from '../../common/styles/global.module.scss';

const LOG_SOURCE: string = "ProjectInformationApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProjectInformationApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ProjectInformationApplicationCustomizer extends BaseApplicationCustomizer<
  IProjectInformationApplicationCustomizerProperties
> {
  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    const body = document.querySelector('body');
    body.classList.add(styles.ELCAVN);

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
      
      // Force to load font-face
      const documentationIcon = React.createElement(Icon, { iconName: 'Documentation'});
      ReactDOM.render(documentationIcon, this._topPlaceholder.domElement);
      
      if (NavService.isProjectPage()) {
        const PI = React.createElement(ProjectInformation, { context: this.context });
        ReactDOM.render(PI, this._topPlaceholder.domElement);
      }
    }
  }

  private _onDispose(): void {
    
  }
}
