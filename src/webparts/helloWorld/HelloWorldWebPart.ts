/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

//import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IKontaktyListItems } from './IKontaktyListItem';
import * as pnp from 'sp-pnp-js';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public onInit(): Promise<void> {
  
    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });

    });    
  }

  public render(): void {
    this.domElement.innerHTML = `<div>
    <div>
    <table border='5' bgcolor='aqua' >


      <tr>
        <td> Podaj id elementu który chcesz wyświetlić </td>
        <td><input type='text' id='txtID' />
        <td><input type='submit' id='btnRead' value='Wyświetl rekord' />
      </tr>

      <tr>
        <td>Name</td>
        <td><input type='text' id='txtName' />
      </tr>
      <tr>
        <td>Surname</td>
        <td><input type='text' id='txtSurname' />
      </tr>
      <tr>
        <td>Phone Number</td>
        <td><input type='text' id='txtPhoneNumber' />
      </tr>

      <tr>
      <td colspan='2' align='center'>
        <input type='submit' value='Dodaj przedmiot' id='btnSubmit' />
        <input type='submit' value='Zmien' id='btnUpdate' />
        <input type='submit' value='Usun' id='btnDelete' />
      </td>
    </table>
    </div>
    <div id="divStatus"/>
    
    </div>`;
      this._bindEvents();
  }

  private _bindEvents(): void {
    this.domElement.querySelector('#btnSubmit')?.addEventListener('click', () => {this.addListItem(); });
    this.domElement.querySelector('#btnRead')?.addEventListener('click', () => {this.readListItem(); });
    this.domElement.querySelector('#btnUpdate')?.addEventListener('click', () => {this.updateListItem();});
    this.domElement.querySelector('#btnDelete')?.addEventListener('click', () => {this.deleteListItem();});
  }
  private deleteListItem(): void {
    const id = parseInt((document.getElementById("txtID") as HTMLInputElement).value);
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    pnp.sp.web.lists.getByTitle("Kontakty").items.getById(id).delete();
    alert("element listy usuniety");
  }
  private updateListItem(): void {
    let Name = (document.getElementById("txtName") as HTMLInputElement).value;
    let Surname = (document.getElementById("txtSurname") as HTMLInputElement).value;
    let PhoneNumber = (document.getElementById("txtPhoneNumber") as HTMLInputElement).value;
    let ID: string = (document.getElementById("txtID") as HTMLInputElement).value;

    
    const siteurl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/Lists/getbytitle('Kontakty')/Items(${ID})`;

    const itemBody: any = {
      "Name": Name,
      "Surname": Surname,
      "PhoneNumber": PhoneNumber
    };
    const headers: any ={
      "X-HTTP-Method" : "MERGE",
      "IF-MATCH" : "*",
    }

    const spHttpClientOptions: ISPHttpClientOptions = {
      "headers": headers,
      "body" : JSON.stringify(itemBody)
    };

    // eslint-disable-next-line no-void
    void this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      response.json().then((responseJson) => {
        console.log(responseJson); // Wyświetl treść odpowiedzi JSON
      }).catch((error) => {
        console.log("blad parsowania odpowiedzi JSON");
      });
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      const statusmessage: Element = this.domElement.querySelector('#divStatus')!;

      if (response.status === 204) {
        statusmessage.innerHTML = "Element został zmieniony pomyslnie.";
        this.clear();
      } else {
        statusmessage.innerHTML = "Blad przy zmianie elementu"
      }
      
    });

  }
  private readListItem(): void {
    let ID: string = (document.getElementById("txtID") as HTMLInputElement).value;
    this._getListItemByID(ID).then(listItem => {
      (document.getElementById("txtName")as HTMLInputElement).value = listItem.Name;
      (document.getElementById("txtSurname")as HTMLInputElement).value = listItem.Surname;
      (document.getElementById("txtPhoneNumber")as HTMLInputElement).value = listItem.PhoneNumber;

    })
    .catch(error => {
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      let message: Element = this.domElement.querySelector('#divStatus')!;
      message.innerHTML = "Read: Could not fetch details... " + error.message;
    });
    
  }
   private _getListItemByID(ID: string): Promise<IKontaktyListItems> {
    const filterValue: string = `$filter=ID eq ${ID}`;
    const url: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/Lists/getbytitle('Kontakty')/items?${filterValue}`;
    return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    .then((Response: SPHttpClientResponse) => {
      
      return Response.json();
      console.log(Response.json);
    })
    .then( (listItems: any)=> {
      
      const untypedItem: any = listItems.value[0];
      const listItem: IKontaktyListItems = untypedItem as IKontaktyListItems;
      return listItem
    }) as Promise <IKontaktyListItems>;
   }
  private addListItem(): void {
    let Name = (document.getElementById("txtName") as HTMLInputElement).value;
    let Surname = (document.getElementById("txtSurname") as HTMLInputElement).value;
    let PhoneNumber = (document.getElementById("txtPhoneNumber") as HTMLInputElement).value;

    const siteurl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/Lists/getbytitle('Kontakty')/items";
    console.log(this.context.pageContext.web.absoluteUrl);

    const itemBody: any = {
      "Name": Name,
      "Surname": Surname,
      "PhoneNumber": PhoneNumber
    };

    const spHttpClientOptions: ISPHttpClientOptions = {
      "body" : JSON.stringify(itemBody)
    };
    console.log(typeof spHttpClientOptions);
    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      response.json().then((responseJson) => {
        console.log(responseJson); // Wyświetl treść odpowiedzi JSON
      }).catch((error) => {
        console.log("blad parsowania odpowiedzi JSON");
      });
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      const statusmessage: Element = this.domElement.querySelector('#divStatus')!;


      if (response.status === 201) {
        statusmessage.innerHTML = "Element został dodany pomyslnie.";
        this.clear();
      } else {
        statusmessage.innerHTML = "Blad przy dodawaniu elementu"
      }
      
    });
  }
  private clear(): void {
    (document.getElementById("txtJobName") as HTMLInputElement).value = '';
    (document.getElementById("txtName") as HTMLInputElement).value = '';
    (document.getElementById("txtSurname") as HTMLInputElement).value = '';
    (document.getElementById("txtPhoneNumber") as HTMLInputElement).value = '';
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
