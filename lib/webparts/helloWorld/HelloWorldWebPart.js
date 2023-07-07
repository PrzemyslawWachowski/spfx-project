var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';
//import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { SPHttpClient } from '@microsoft/sp-http';
import * as pnp from 'sp-pnp-js';
var HelloWorldWebPart = /** @class */ (function (_super) {
    __extends(HelloWorldWebPart, _super);
    function HelloWorldWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    HelloWorldWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            pnp.setup({
                spfxContext: _this.context
            });
        });
    };
    HelloWorldWebPart.prototype.render = function () {
        this.domElement.innerHTML = "<div>\n    <div>\n    <table border='5' bgcolor='aqua' >\n\n\n      <tr>\n        <td> Podaj id elementu kt\u00F3ry chcesz wy\u015Bwietli\u0107 </td>\n        <td><input type='text' id='txtID' />\n        <td><input type='submit' id='btnRead' value='Wy\u015Bwietl rekord' />\n      </tr>\n\n      <tr>\n        <td>Name</td>\n        <td><input type='text' id='txtName' />\n      </tr>\n      <tr>\n        <td>Surname</td>\n        <td><input type='text' id='txtSurname' />\n      </tr>\n      <tr>\n        <td>Phone Number</td>\n        <td><input type='text' id='txtPhoneNumber' />\n      </tr>\n\n      <tr>\n      <td colspan='2' align='center'>\n        <input type='submit' value='Dodaj przedmiot' id='btnSubmit' />\n        <input type='submit' value='Zmien' id='btnUpdate' />\n        <input type='submit' value='Usun' id='btnDelete' />\n      </td>\n    </table>\n    </div>\n    <div id=\"divStatus\"/>\n    \n    </div>";
        this._bindEvents();
    };
    HelloWorldWebPart.prototype._bindEvents = function () {
        var _this = this;
        var _a, _b, _c, _d;
        (_a = this.domElement.querySelector('#btnSubmit')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', function () { _this.addListItem(); });
        (_b = this.domElement.querySelector('#btnRead')) === null || _b === void 0 ? void 0 : _b.addEventListener('click', function () { _this.readListItem(); });
        (_c = this.domElement.querySelector('#btnUpdate')) === null || _c === void 0 ? void 0 : _c.addEventListener('click', function () { _this.updateListItem(); });
        (_d = this.domElement.querySelector('#btnDelete')) === null || _d === void 0 ? void 0 : _d.addEventListener('click', function () { _this.deleteListItem(); });
    };
    HelloWorldWebPart.prototype.deleteListItem = function () {
        var id = parseInt(document.getElementById("txtID").value);
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        pnp.sp.web.lists.getByTitle("Kontakty").items.getById(id).delete();
        alert("element listy usuniety");
    };
    HelloWorldWebPart.prototype.updateListItem = function () {
        var _this = this;
        var Name = document.getElementById("txtName").value;
        var Surname = document.getElementById("txtSurname").value;
        var PhoneNumber = document.getElementById("txtPhoneNumber").value;
        var ID = document.getElementById("txtID").value;
        var siteurl = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/Lists/getbytitle('Kontakty')/Items(").concat(ID, ")");
        var itemBody = {
            "Name": Name,
            "Surname": Surname,
            "PhoneNumber": PhoneNumber
        };
        var headers = {
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
        };
        var spHttpClientOptions = {
            "headers": headers,
            "body": JSON.stringify(itemBody)
        };
        // eslint-disable-next-line no-void
        void this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then(function (response) {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            response.json().then(function (responseJson) {
                console.log(responseJson); // Wyświetl treść odpowiedzi JSON
            }).catch(function (error) {
                console.log("blad parsowania odpowiedzi JSON");
            });
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            var statusmessage = _this.domElement.querySelector('#divStatus');
            if (response.status === 204) {
                statusmessage.innerHTML = "Element został zmieniony pomyslnie.";
                _this.clear();
            }
            else {
                statusmessage.innerHTML = "Blad przy zmianie elementu";
            }
        });
    };
    HelloWorldWebPart.prototype.readListItem = function () {
        var _this = this;
        var ID = document.getElementById("txtID").value;
        this._getListItemByID(ID).then(function (listItem) {
            document.getElementById("txtName").value = listItem.Name;
            document.getElementById("txtSurname").value = listItem.Surname;
            document.getElementById("txtPhoneNumber").value = listItem.PhoneNumber;
        })
            .catch(function (error) {
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            var message = _this.domElement.querySelector('#divStatus');
            message.innerHTML = "Read: Could not fetch details... " + error.message;
        });
    };
    HelloWorldWebPart.prototype._getListItemByID = function (ID) {
        var filterValue = "$filter=ID eq ".concat(ID);
        var url = "".concat(this.context.pageContext.web.absoluteUrl, "/_api/web/Lists/getbytitle('Kontakty')/items?").concat(filterValue);
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then(function (Response) {
            return Response.json();
            console.log(Response.json);
        })
            .then(function (listItems) {
            var untypedItem = listItems.value[0];
            var listItem = untypedItem;
            return listItem;
        });
    };
    HelloWorldWebPart.prototype.addListItem = function () {
        var _this = this;
        var Name = document.getElementById("txtName").value;
        var Surname = document.getElementById("txtSurname").value;
        var PhoneNumber = document.getElementById("txtPhoneNumber").value;
        var siteurl = this.context.pageContext.web.absoluteUrl + "/_api/web/Lists/getbytitle('Kontakty')/items";
        console.log(this.context.pageContext.web.absoluteUrl);
        var itemBody = {
            "Name": Name,
            "Surname": Surname,
            "PhoneNumber": PhoneNumber
        };
        var spHttpClientOptions = {
            "body": JSON.stringify(itemBody)
        };
        console.log(typeof spHttpClientOptions);
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then(function (response) {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            response.json().then(function (responseJson) {
                console.log(responseJson); // Wyświetl treść odpowiedzi JSON
            }).catch(function (error) {
                console.log("blad parsowania odpowiedzi JSON");
            });
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            var statusmessage = _this.domElement.querySelector('#divStatus');
            if (response.status === 201) {
                statusmessage.innerHTML = "Element został dodany pomyslnie.";
                _this.clear();
            }
            else {
                statusmessage.innerHTML = "Blad przy dodawaniu elementu";
            }
        });
    };
    HelloWorldWebPart.prototype.clear = function () {
        document.getElementById("txtJobName").value = '';
        document.getElementById("txtName").value = '';
        document.getElementById("txtSurname").value = '';
        document.getElementById("txtPhoneNumber").value = '';
    };
    Object.defineProperty(HelloWorldWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    HelloWorldWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return HelloWorldWebPart;
}(BaseClientSideWebPart));
export default HelloWorldWebPart;
//# sourceMappingURL=HelloWorldWebPart.js.map