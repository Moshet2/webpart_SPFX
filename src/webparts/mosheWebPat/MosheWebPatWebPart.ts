import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MosheWebPatWebPart.module.scss';
import * as strings from 'MosheWebPatWebPartStrings';
 
export interface IMosheWebPatWebPartProps {
  description: string;  
}
 
import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import { isNull } from 'lodash';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  TZ : string;
  Job :string;
}
   
 export default class MosheWebPatWebPart extends BaseClientSideWebPart<IMosheWebPatWebPartProps>{
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _severTest: string = "https://6jhgyz.sharepoint.com/sites/Test";

  readonly absoluteUrl: string; 
    private _getListData(): Promise<ISPLists> {
      
  //  return this.context.spHttpClient.get("https://6jhgyz.sharepoint.com/sites/Test/_api/web/lists/GetByTitle('Customes')/Items",SPHttpClient.configurations.v1)      
  //   return this.context.spHttpClient.get( this.context.pageContext.web.absoluteUrl +"/_api/web/lists/GetByTitle('Customes')/Items",SPHttpClient.configurations.v1)
  return this.context.spHttpClient.get( this._severTest +"/_api/web/lists/GetByTitle('Customes')/Items",SPHttpClient.configurations.v1)

        .then((response: SPHttpClientResponse) => {
        return response.json();
        });
    }
    private _renderListAsync(): void {
    
      if (Environment.type == EnvironmentType.SharePoint || 
               Environment.type == EnvironmentType.ClassicSharePoint) {
       this._getListData()
         .then((response) => {
           this._renderList(response.value);
         });
     }
   }
    private _renderList(items: ISPList[]): void {
      let html: string = '<table border=1 width=100% style="border-collapse: collapse;">';
      html += '<th>Name</th> <th>TZ</th><th>Job</th>';
      items.forEach((item: ISPList) => {
        html += `
        <tr>            
        <td>${item.Title}</td>
        <td>${item.TZ}</td>
        <td>${item.Job}</td> 
            
            </tr>
            `;
      });
      html += '</table>';
    
      const listContainer: Element = this.domElement.querySelector('#spListContainer');
      listContainer.innerHTML = html;
    }
    
    

    public render(): void {
      this.domElement.innerHTML = `
        <div class="${ styles.mosheWebPat } ">
          <div class="${ styles.container } ">
            <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${ styles.row }">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint Modern Web Part By Moshe Test</span>
            <p class="ms-font-l ms-fontColor-white">Loading from ${this.context.pageContext.web.title} Web Site</p>
            <p class="ms-font-l ms-fontColor-white">List Data from SharePoint List Customers </p>
          </div>
        </div> 
            <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${ styles.row }">
            <div>List Items</div>
            <br>
             <div id="spListContainer" />
          </div>
          <div class="${ styles.container2 }   ">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${ styles.row }">
          <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <span class="ms-font-xl ms-fontColor-white">Create Customer By Moshe Test</span> 
        </div>
        </div> 
        <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${ styles.row }">
        <div>Customer Items</div>
        <br>
         <div id="spListContainer2" />
            <div class="col-sm-6">
            <label for="Name" >Input Name: </label>
            <input type="text" id="name" name="Name" ><br><br>
            </div>
            <div class="col-sm-6">
            <label for="Name" >Input TZ: </label>
            <input type="text" id="tz" name="TZ" ><br><br>
            </div>
            <div class="col-sm-6">
            <label for="Name" >Input Job: </label>
            <input type="text" id="job" name="Job" ><br><br>
            </div>
            <div class="${ styles.Savebutton } "> 
            <input type="button" id="saveButton"  value="Save" ><br><br>
            </div>
      </div>
        </div>`;

         this._bindSave();
         this._renderListAsync();
    }
    
    private _bindSave(): void{
       this.domElement.querySelector("#saveButton").addEventListener('click', () => {this.addListItem();
      });
     // alert("button Save");
    }

    private addListItem(): void{
      var flag= true;
      var Name = document.getElementById("name")["value"];
      var TZ = document.getElementById("tz")["value"];
      var Job = document.getElementById("job")["value"];
       
      if (Name == ''|| TZ == ''|| Job =='' ){
          alert("Customer popertis is not full");
          flag = false;
      }

      if (flag){
        //  const siteUrl: string = this.context.pageContext.site.absoluteUrl +"_api/web/lists/GetByTitle('Customes')/Items";
     const siteUrl: string = this._severTest + "_api/web/lists/GetByTitle('Customes')/Items";
      //  const siteUrl: string = "https://6jhgyz.sharepoint.com/sites/Test/_api/web/lists/GetByTitle('Customes')/Items";
        const itemBody: any={
          "Title": Name,
          "TZ": TZ,
          "Job": Job
        }
     
     const spHttpClientOptions: ISPHttpClientOptions ={
       "body": JSON.stringify(itemBody)
     }; 
    
     this.context.spHttpClient.post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
     .then((response: SPHttpClientResponse) => {
       alert('Customer Saved');
       this._renderListAsync();
       document.getElementById("name")["value"]="";
      document.getElementById("tz")["value"]="";
      document.getElementById("job")["value"]="";
      });
      }
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