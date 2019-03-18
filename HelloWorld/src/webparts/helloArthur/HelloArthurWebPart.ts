import { 
  Environment,
  EnvironmentType,
  Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import MockHttpClient from './MockHttpClient';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import styles from './HelloArthurWebPart.module.scss';
import * as strings from 'HelloArthurWebPartStrings';

export interface IHelloArthurWebPartProps {
  character: string;
  description: string;
  motto: string;
  isPanicking: boolean;
  equippedItem: string;
  hasTea: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloArthurWebPart extends BaseClientSideWebPart<IHelloArthurWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloArthur }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${this.properties.character}</span>
              <p class="${ styles.subTitle }"><em>"${escape(this.properties.motto)}"</em></p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <p class="${ styles.description }">${this.properties.character} is holding: ${this.properties.equippedItem},</p>
              <p class="${ styles.description }">The thermos is currently ${this.properties.hasTea.valueOf() ? 'full' : 'empty' }</p>
              <p class="${ styles.description }">and right now s/he is ${this.properties.isPanicking.valueOf() ? 'panicking' : 'calm' }.</p>
            </div>
          </div>
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <p><b>Available Equipment:</b></p>
              <div id="spListContainer" />
              <p></p>
              <p class="">Webpart loaded from ${escape(this.context.pageContext.web.title)}</p>
            </div>
          </div>         
        </div>
      </div>`;

      this._renderListAsync();
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
                PropertyPaneTextField('character', {
                  label: 'Who are we talking about?'
                }),
                PropertyPaneTextField('description', {
                  label: 'A description of the character.',
                  multiline: true
                }),
                PropertyPaneTextField('motto', {
                  label: 'The character\s life motto.'
                }),
                PropertyPaneCheckbox('isPanicking',{
                  text: 'Whether or not this character is freaking out.'
                }),
                PropertyPaneDropdown('equippedItem', {
                  label: 'Equip an item.',
                  options: [
                    {key: 'A thermos flask', text: 'A thermos flask'},
                    {key: 'A towel', text: 'A towel'},
                    {key: 'Some aspirin', text: 'Some aspirin'},
                    {key: 'The Hitch Hiker\'s Guide to the Galaxy', text: 'The Hitch Hiker\'s Guide to the Galaxy'}
                  ]
                }),
                PropertyPaneToggle('hasTea', {
                  label: 'Toggle whether or not the thermos has tea:',
                  onText: 'Full/Has Tea',
                  offText: 'Empty/No Tea'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _renderListAsync(): void {
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type === EnvironmentType.SharePoint
            || Environment.type === EnvironmentType.ClassicSharePoint)
    {
      this._getListData().then((response) => {
        this._renderList(response.value);
      });
    }
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '<ul class="${styles.list}">';
    items.forEach((item: ISPList) => {
      html += `
        <li class="${styles.listItem}">
          <span class="ms-font-l">${item.Title}</span>
        </li>`;
    });
    html += "</ul>";
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = (html);
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = {value: data };
        return listData;
      }) as Promise<ISPLists>;
  }


}
