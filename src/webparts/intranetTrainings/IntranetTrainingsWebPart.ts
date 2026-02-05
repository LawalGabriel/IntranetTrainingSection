import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown
  
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'IntranetTrainingsWebPartStrings';
import IntranetTrainings from './components/IntranetTrainings';
import { IIntranetTrainingsProps } from './components/IIntranetTrainingsProps';

export interface IIntranetTrainingsWebPartProps {
  enableScroll: boolean;
  dateBackgroundColor: string;
  maxRowsBeforeScroll: number;
  categoryBgColor: string;
  categoryColor: string;
  showCategory: boolean;
  cardHeight: number;
  description: string;
  listTitle: string;
  webPartTitle: string;
  cardBackgroundColor: string;
  cardBorderColor: string;
  titleColor: string;
  dateColor: string;
  useFullWidth: boolean;
  itemsPerRow: number;
  showScheduledDate: boolean;
  titleFontColor: string;
  titleFontWeight: string;
  titleFontSize: number;
  titleBackgroundColor: string;
}

export default class IntranetTrainingsWebPart extends BaseClientSideWebPart<IIntranetTrainingsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IIntranetTrainingsProps> = React.createElement(
      IntranetTrainings,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        // Pass all properties from this.properties
        listTitle: this.properties.listTitle || "TrainingRepo",
        context: this.context,
        webPartTitle: this.properties.webPartTitle || "Training Items",
        cardBackgroundColor: this.properties.cardBackgroundColor || "#ffffff",
        cardBorderColor: this.properties.cardBorderColor || "#e1e1e1",
        titleColor: this.properties.titleColor || "#000000",
        dateColor: this.properties.dateColor || "#333333",
        useFullWidth: this.properties.useFullWidth || false,
        itemsPerRow: this.properties.itemsPerRow || 3,
        showScheduledDate: this.properties.showScheduledDate || false,
        // New web part title properties
        titleFontColor: this.properties.titleFontColor || "#000000",
        titleFontWeight: this.properties.titleFontWeight || "600",
        titleFontSize: this.properties.titleFontSize || 32,
        titleBackgroundColor: this.properties.titleBackgroundColor || "transparent",
      // In the render method where you pass props:
cardHeight: Math.max(this.properties.cardHeight || 80, 0),
        // Category properties
        showCategory: this.properties.showCategory || false,
        categoryColor: this.properties.categoryColor || "#333333",
        categoryBgColor: this.properties.categoryBgColor || "#f0f0f0",
        maxRowsBeforeScroll: this.properties.maxRowsBeforeScroll || 5,
        dateBackgroundColor: this.properties.dateBackgroundColor || "#f8f9fa",
        enableScroll: this.properties.enableScroll || false
      
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
              PropertyPaneTextField('listTitle', {
                label: 'List Title',
                value: this.properties.listTitle || "TrainingRepo"
              }),
              PropertyPaneToggle('useFullWidth', {
                label: 'Use Full Width',
                onText: 'Yes',
                offText: 'No',
                checked: this.properties.useFullWidth || false
              }),
              PropertyPaneSlider('itemsPerRow', {
                label: 'Items per Row',
                min: 1,
                max: 4,
                value: this.properties.itemsPerRow || 3,
                showValue: true
              }),
              PropertyPaneSlider('cardHeight', {
                label: 'Card Height (px)',
                min: 0,
                max: 100,
                value: this.properties.cardHeight || 80,
                showValue: true,
                step: 5
              }),
              PropertyPaneToggle('showScheduledDate', {
                label: 'Show Scheduled Date',
                onText: 'Show',
                offText: 'Hide',
                checked: this.properties.showScheduledDate || false
              }),
              PropertyPaneToggle('showCategory', {
                label: 'Show Category',
                onText: 'Show',
                offText: 'Hide',
                checked: this.properties.showCategory || false
              })
            ]
          },
          {
            groupName: 'Scroll Settings',
            groupFields: [
              PropertyPaneToggle('enableScroll', {
                label: 'Enable Scroll',
                onText: 'Yes',
                offText: 'No',
                checked: this.properties.enableScroll || false,
                //description: 'Enable scroll when items exceed maximum rows'
              }),
              PropertyPaneSlider('maxRowsBeforeScroll', {
                label: 'Max Rows Before Scroll',
                min: 1,
                max: 10,
                value: this.properties.maxRowsBeforeScroll || 5,
                showValue: true,
                step: 1
              })
            ]
          },
          {
            groupName: 'Web Part Title Settings',
            groupFields: [
              PropertyPaneTextField('webPartTitle', {
                label: 'Title Text',
                value: this.properties.webPartTitle || "Training Items",
                description: 'Enter the title for the web part'
              }),
              PropertyPaneTextField('titleFontColor', {
                label: 'Title Font Color',
                value: this.properties.titleFontColor || "#000000",
                description: 'Enter color in hex format (e.g., #000000)'
              }),
              PropertyPaneDropdown('titleFontWeight', {
                label: 'Title Font Weight',
                selectedKey: this.properties.titleFontWeight || "600",
                options: [
                  { key: '300', text: 'Light (300)' },
                  { key: '400', text: 'Normal (400)' },
                  { key: '500', text: 'Medium (500)' },
                  { key: '600', text: 'Semi-bold (600)' },
                  { key: '700', text: 'Bold (700)' },
                  { key: '800', text: 'Extra-bold (800)' }
                ]
              }),
              PropertyPaneSlider('titleFontSize', {
                label: 'Title Font Size (px)',
                min: 16,
                max: 48,
                value: this.properties.titleFontSize || 32,
                showValue: true,
                step: 2
              }),
              PropertyPaneTextField('titleBackgroundColor', {
                label: 'Title Background Color',
                value: this.properties.titleBackgroundColor || "transparent",
                description: 'Enter color in hex format or "transparent"'
              })
            ]
          },
          {
            groupName: 'Card Design Settings',
            groupFields: [
              PropertyPaneTextField('cardBackgroundColor', {
                label: 'Card Background Color',
                value: this.properties.cardBackgroundColor || "#ffffff",
                description: 'Enter color in hex format'
              }),
              PropertyPaneTextField('cardBorderColor', {
                label: 'Card Border Color',
                value: this.properties.cardBorderColor || "#e1e1e1",
                description: 'Enter color in hex format'
              }),
              PropertyPaneTextField('titleColor', {
                label: 'Item Title Color',
                value: this.properties.titleColor || "#000000",
                description: 'Enter color in hex format'
              })
            ]
          },
          {
            groupName: 'Date Section Settings',
            groupFields: [
              PropertyPaneTextField('dateColor', {
                label: 'Date Text Color',
                value: this.properties.dateColor || "#333333",
                description: 'Enter color in hex format'
              }),
              PropertyPaneTextField('dateBackgroundColor', {
                label: 'Date Background Color',
                value: this.properties.dateBackgroundColor || "#f8f9fa",
                description: 'Enter color in hex format'
              })
            ]
          },
          {
            groupName: 'Category Settings',
            groupFields: [
              PropertyPaneTextField('categoryColor', {
                label: 'Category Text Color',
                value: this.properties.categoryColor || "#333333",
                description: 'Enter color in hex format'
              }),
              PropertyPaneTextField('categoryBgColor', {
                label: 'Category Background Color',
                value: this.properties.categoryBgColor || "#f0f0f0",
                description: 'Enter color in hex format'
              })
            ]
          }
        ]
      }
    ]
  };
}
}