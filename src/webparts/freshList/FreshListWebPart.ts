import * as React from 'react';
import * as ReactDom from 'react-dom';
import { sp } from '@pnp/sp';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FreshListWebPartStrings';
import FreshList from './components/FreshList';
import { IFreshListProps } from './components/IFreshListProps';

export interface IFreshListWebPartProps {
  description: string;
  listUrlAPI: string;
  numberOfElement: any;
  themeColor: any;
  webPartName: string;
  columnSelected: string;
  listName: string;
  themeID: number;
}

export default class FreshListWebPart extends BaseClientSideWebPart<IFreshListWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IFreshListProps> = React.createElement(
      FreshList,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        listUrlAPI: this.properties.listUrlAPI,
        numberOfElement: this.properties.numberOfElement,
        themeColor: this.properties.themeColor,
        webPartName: this.properties.webPartName,
        columnSelected: this.properties.columnSelected,
        listName: this.properties.listName,
        context: this.context,
        themeID: this.properties.themeID,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private Call = (e) => {
    alert("Button 1, Clicked");
  } 

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Fresh crud webpart"
          },

          groups: [

            // config of webpart name
            {
              groupName: 'Configuration du webPart',
              groupFields: [
                PropertyPaneTextField('webPartName', {
                  label: "Entrer le nom de votre webPart"
                })
              ]
            },

            // config of list (API, number of element selected and items selected)
            {
              groupName: "Configuration de la liste",
              groupFields: [
                PropertyPaneTextField('listUrlAPI', {
                  label: "Entrer l'API de votre liste"
                }),
                PropertyPaneSlider('numberOfElement', { 
                  label: 'Entrer le maximum des elements', 
                  min: 1, 
                  max: 20, 
                  showValue: true, 
                  value: 10 
                }),
                PropertyPaneTextField('columnSelected', {
                  label: "Entrer les colonnes sélectionné"
                }),
              ]
            },
            {
              groupName: "Configuration de l'ajout / modif / supp",
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: "Entrer le nom de votre liste"
                }),
              ]
            },

            
            
            
          ]
          
        },
        {
          header: {
            description: "Fresh crud webpart"
          },

          groups: [

            // config of webpart name
            {
              groupName: 'Configuration du theme',
              groupFields: [
                PropertyPaneDropdown('themeID',{
                  label: 'Choisir votre theme',
                  selectedKey: 1,
                  disabled: false,
                  options:[
                    {
                      key:1,
                      text:"White theme"
                    },
                    {
                      key:2,
                      text:"Primary theme"
                    },
                    {
                      key:3,
                      text:"black theme"
                    },
                    {
                      key:4,
                      text:"Secondary theme"
                    }
                  ]

                })
              ]
            },

            
            
            
            
            
          ]
          
        }

        
      ]
    };
  }
}
