import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './SpfxContextExplorerWebPart.module.scss';
import * as strings from 'mystrings';
import { ISpfxContextExplorerWebPartWebPartProps } from './ISpfxContextExplorerWebPartWebPartProps';

import { EnvironmentType } from '@microsoft/sp-client-base';

export default class SpfxContextExplorerWebPartWebPart extends BaseClientSideWebPart<ISpfxContextExplorerWebPartWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.spfxContextExplorerWebPart}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Wes Hackett - SPFx Learning - Context Explorer</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
              <a href="https://github.com/SharePoint/sp-dev-docs/wiki" class="ms-Button ${styles.button}">
                <span class="ms-Button-label">Learn more</span>
              </a>

              <h3>context.instanceId</h3>
              <p>The unique web part Id</p>
              <ul>
                 <li>instanceId: ${this.context.instanceId}</li>
              </ul>

              <h3>context.environment</h3>
              <p>EnvironmentType module</p>
              <p>SharePoint Workbench gives you the flexibility to test web parts in your local environment and from a SharePoint site. SharePoint Framework aids this capability by helping you understand which environment your web part is running from using the  EnvironmentType  module.</p>
              <p>Import: import { EnvironmentType } from '@microsoft/sp-client-base';</p>
              <ul>
                 <li>name (converted): ${this._getEnvironmentTypeName()}</li>
                 <li>type: ${this.context.environment.type}</li>
              </ul>

              <h3>context.manifest</h3>
              <p>The web part manifest information.</p>
              <ul>
                 <li>componentType: ${this.context.manifest.componentType}</li>
                 <li>description.default: ${this.context.manifest.description.default}</li>
                 <li>group.default: ${this.context.manifest.group.default}</li>
                 <li>groupId: ${this.context.manifest.groupId}</li>
                 <li>iconImageUrl: ${this.context.manifest.iconImageUrl}</li>
                 <li>id: ${this.context.manifest.id}</li>
                 <li>imageLinkPropertyNames: ${this.context.manifest.imageLinkPropertyNames}</li>
                 <li>linkPropertyNames: ${this.context.manifest.linkPropertyNames}</li>
                 <li>context.manifest.loaderConfig
                    <ul>
                      <li>loaderConfig.entryModuleId: ${this.context.manifest.loaderConfig.entryModuleId}</li>
                      <li>loaderConfig.exportedModuleName: ${this.context.manifest.loaderConfig.exportedModuleName}</li>
                      <li>loaderConfig.internalModuleBaseUrls: ${this.context.manifest.loaderConfig.internalModuleBaseUrls}</li>
                      <li>loaderConfig.scriptResources: ${this.context.manifest.loaderConfig.scriptResources}</li>
                    </ul>
                 <li>manifestVersion: ${this.context.manifest.manifestVersion}</li>
                 <li>officeFabricIconFontName: ${this.context.manifest.officeFabricIconFontName}</li>
                 <li>properties: ${this._getManifestProperties()}</li>
                 <li>searchablePropertyNames: ${this.context.manifest.searchablePropertyNames}</li>
                 <li>title.default: ${this.context.manifest.title.default}</li>
                 <li>version: ${this.context.manifest.version}</li>
              </ul>

              <h3>context.pageContext</h3>
              <p>The SharePoint page context.</p>
              <ul>
                  <li>pageContext.cultureInfo
                    <ul>
                      <li>currentCultureName: ${this.context.pageContext.cultureInfo.currentCultureName}</li>
                      <li>currentUICultureName: ${this.context.pageContext.cultureInfo.currentUICultureName}</li>
                    </ul>
                  </li>
                  <li>isInitialized: ${this.context.pageContext.isInitialized}</li>
                  <li>pageContext.site
                    <ul>
                      <li>id: ${this.context.pageContext.site.id}</li>
                    </ul>
                  </li>
                  <li>pageContext.user
                    <ul>
                      <li>displayName: ${this.context.pageContext.user.displayName}</li>
                      <li>loginName: ${this.context.pageContext.user.loginName}</li>
                    </ul>
                  </li>
                  <li>pageContext.web
                    <ul>
                      <li>absoluteUrl: ${this.context.pageContext.web.absoluteUrl}</li>
                      <li>id: ${this.context.pageContext.web.id}</li>
                      <li>serverRelativeUrl: ${this.context.pageContext.web.serverRelativeUrl}</li>
                      <li>title: ${this.context.pageContext.web.title}</li>
                    </ul>
                  </li>
              </ul>

              <h3>context.webPartTag</h3>
              <p>The web part tag to use for telementry</p>
              <p>Format appears to be {manifest.componentType}.{manifest.id}.{instanceId}
              <ul>
                 <li>webPartTag: ${this.context.webPartTag}</li>
              </ul>

            </div>
          </div>
        </div>
      </div>`;
  }

  private _getEnvironmentTypeName() : string {

    let value: string;
    switch (this.context.environment.type) {
      case EnvironmentType.ClassicSharePoint:
        value = 'ClassicSharePoint';
        break;
      case EnvironmentType.Local:
        value = 'Local';
        break;
      case EnvironmentType.SharePoint:
        value = 'SharePoint';
        break;
      case EnvironmentType.Test:
        value = 'Test';
        break;
      default:
        value = 'Unknown';
        break;
    }

    return value;
  }

  private _getManifestProperties() : string {
    let props: string;
    props = '<ul>';

    for (var index = 0; index < this.context.manifest.properties.length; index++) {
      var element = this.context.manifest.properties[index];
      props = props + '<li>' + element + '</li>';
    }


    props = props + '</ul>';

    return props;
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
