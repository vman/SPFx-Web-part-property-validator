import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import styles from './PropValidator.module.scss';
import * as strings from 'mystrings';
import { IPropValidatorWebPartProps } from './IPropValidatorWebPartProps';

export default class PropValidatorWebPart extends BaseClientSideWebPart<IPropValidatorWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.propValidator}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">${this.properties.title}</span>
              <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
            </div>
          </div>
        </div>
      </div>`;
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  onGetErrorMessage: this._validateTitleAsync.bind(this), // validation function
                  deferredValidationTime: 500 // delay after which to run the validation function
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  resizable: true,
                  onGetErrorMessage: this._validateDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _validateTitleAsync(value: string): Promise<string> {

    return this.context.httpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/title`)
      .then((response: Response) => {
        return response.json().then((responseJSON) => {
          if (responseJSON.value.toLowerCase() === value.toLowerCase()) {
            return "Title cannot be the same as the SharePoint site title";
          }
          else {
            return "";
          }
        });
      });

  }

  private _validateDescription(value: string): string {
    if (value.length < 10) {
      return "At least 10 characters required";
    }
    else {
      return "";
    }
  }
}