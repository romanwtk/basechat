import { Version } from '@microsoft/sp-core-library';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

import styles from './RwBaseChatWebPart.module.scss';
import * as strings from 'RwBaseChatWebPartStrings';
import jQuery from 'jquery';

// Import der Bibliothek PnPjs - hier in Verwendung für REST-Calls und Listenabfragen
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import pnp, { List, ListEnsureResult } from "sp-pnp-js";

// Benötigte Imports für HTTP-Requests
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

export interface IRwBaseChatWebPartProps {
  list: string;
  title: string;
  background: string;
  allowhyperlinks: boolean;
  moderators: IPropertyFieldGroupOrPerson[];
  interval: any;
  autofetch: boolean;
  emojiset: string;
}

export interface ISPLists {
  value: ISPList[];
}

// Definition einer Chatnachricht als Liste
export interface ISPList {
  Id: string;
  Time: string;
  User: string;
  Message: string;
  AnswerTo: string;
  Pinned: string;
}

export interface IField {
  Title: string;
}

// Globale Variablen, die später verwendet werden
let listName: string;
let messages = 0;

export default class RwBaseChatWebPart extends BaseClientSideWebPart <IRwBaseChatWebPartProps> {
  // Benötigt, um PnPjs verwenden zu können
  public 'use strict';

  /*  Die onInit-Methode wird verwendet, um die Voraussetzungen für HTTP-Requests über PnPjs zu ermöglichen. 
      Ich habe sie hier außerdem zweckentfremdet, um setInterval-Aufrufe nutzen zu können, da das in der render-Methode nicht funktioniert.
  */
  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      pnp.setup({
        sp: {
          baseUrl: document.location.href
        }
      });
      sp.setup({
        sp: {
          baseUrl: document.location.href
        }
      });

      // Prüft im im PropertyPane definierten Intervall, ob neue Chat-Nachrichten da sind
      if (this.properties.autofetch == true) {
        setInterval((e) => this.checkChatList(this.properties.list), this.properties.interval);
      }
    });
  }
  
  /*  Logik dahinter: Vergleiche die Anzahl der bereits heruntergeladenen mit der Anzahl der Nachrichten in der Liste.
      Wenn die Zahlen nicht übereinstimmen, werden alle Nachrichten neu gerendert und angzeiegt.

      Optimierbar: Vergleiche nur die IDs der Nachrichten und zeige nur die neuesten Nachrichten zusätzlich an.
      Also anstatt eines vollständigen Austausches nur neue Nachrichten - Problem: Mögliche Löschungen/Zensuren werden nicht mit abgebildet.
  */

  public checkChatList(list): void {
    this._getListData()
    .then((response) => {
      if (response.value.length != messages) {
        this._renderList(response.value);
      }
      else {
        // Do nothing
      }
    });
  }

  // Setzt den Reaktionstyp der PropertyPane auf nicht-rektiv, um Listenerstellungen zu händeln.
  protected get disableReactivePropertyChanges(): boolean { 
    return true; 
  }

  /*  
   *  */  
  public render(): void {

    // Greift den Listennamen aus den Properties und weist ihn der "globalen" (fungiert nur so) Variable listName zu.
    listName = this.properties.list;
    this.domElement.innerHTML = `
      <div class="${ styles.rwBaseChat }">
        <div class="${ styles.container }" >
          <div class="${ styles.row }" style="background-color: ${this.properties.background};">
            <p class="${ styles.title}">${escape(this.properties.title)}<div id="spPinnedContainer"></div></p><br>
            <div id="spListContainer" style="max-height: 300px; overflow: auto; background-color: ${this.properties.background}" />
            <bR>
            <bR>
          </div>
          <div class="${ styles.row }" style="background-color: ${this.properties.background};">
            <table style="width: 100%; text-align: center;">
              <tr>
                <td>
                  <button id="emoji" class="${ styles.button2 }"><span>&#128521;</span></button>
                </td>
                <td>
                  <textarea id="message" placeholder="Deine Nachricht" name="message" class="${styles.textarea}"></textarea>
                </tD>
                <td>
                  <button id="submit" class="${ styles.button2 }"><span>&#9993;</span></button>
                </td>
              </tr>
            </table>
            <div id="emojibar" class="${styles.emojibar}">
              <!-- Alle Emojis sind Unicode-Symbole -->
              <button class="${styles.emoji}">😂</button>
            </div>
          </div>
        </div>
      </div>`;

      let emojiSet = this.properties.emojiset;
      try {
        let emojis = emojiSet.split(";");
        let emojihtml = '';
        emojis.forEach(em => {
          emojihtml += '<button class="'+styles.emoji+'">'+em+'</button>';
        });
        const emojibar: Element = this.domElement.querySelector('#emojibar');
          emojibar.innerHTML = emojihtml;
      }
      catch {
        //
      }
    
    // Bei Klick auf den "Sende"-Button (Briefumschlag) wird die Funktion _getMessage() aufgerufen
    let btnSubmit = this.domElement.querySelector("#submit");
    btnSubmit.addEventListener("click", (e:Event) => this._getMessage());

    // Ruft alle Emoji-Buttons auf. Die Notation im querySelector ist daher so erforderlich, da SharePoint die Klassendefinitionen um IDs erweitert.
    let btnEmojiButton= this.domElement.querySelectorAll("."+styles.emoji);

    // Jedem Button der Klasse emoji wird ein EventListener hinzugefügt.
    btnEmojiButton.forEach((item) => {
      item.addEventListener("click", (e:Event) => this.innerEmoji(item.textContent));
    });

    // Nicht fertiggestellter Ansatz für eine Antwort-Funktion
    let btnAnswerButton = this.domElement.querySelectorAll("."+styles.answer);
    btnAnswerButton.forEach((item) => {
      item.addEventListener("click", (e:Event) => this.setAnswer(item.id));
    });

    // Ermöglicht es, die Emoji-Leiste ein- und auszuklappen
    let btnEmoji = this.domElement.querySelector("#emoji");
    btnEmoji.addEventListener("click", (e:Event) => {
      if (this.domElement.querySelector("#emojibar").classList.contains(styles.emojibar)) {
        this.domElement.querySelector("#emojibar").classList.add(styles.emojibarOpen);
        this.domElement.querySelector("#emojibar").classList.remove(styles.emojibar);
      }
      else {
        this.domElement.querySelector("#emojibar").classList.add(styles.emojibar);
        this.domElement.querySelector("#emojibar").classList.remove(styles.emojibarOpen);
      }
    });
    
    // Initial-Aufruf, um sowohl zu prüfen, ob die Liste existiert.
    this.checkListCreation(this.properties.list);

    // Initial-Aufruf, der die Nachrichten herunterlädt
    this._renderListAsync();
  }
  
  // Fügt das gewählte Emoji in den Text ein
  public innerEmoji(emoji) {
    let txaMessage = this.domElement.querySelector('textarea');
    txaMessage.value += emoji;
  }

  // Ansatz für ie Antwortfunktion - nicht fertiggestellt
  public setAnswer(reference) {
    console.log("You're triggered!");
    let txaMessage = this.domElement.querySelector('textarea');
    txaMessage.value += '<a href="#'+reference+'">Antwort</a>';
  }
  

  public async _getMessage() {
    // Fragt den Anmeldenamen des Users ab. Bei Hauptkonto-Adressen wird der Anzeigename gewählt.
    let user = this.context.pageContext.user.displayName;
    // Bei Gästen wird der Loginname, meistens die E-Mail-Adresse, abgerufen
    // Fungiert hier als Rückfallebene, falls kein Anzeigename existiert
    if (this.context.pageContext.user.isExternalGuestUser == true) {
      user = this.context.pageContext.user.loginName;
    }
    
    // Ruft die Nachricht aus dem Textfeld ab
    let message = this.domElement.querySelector("textarea").value;

    // Ruft die Zeit ab.
    let time = new Date().toLocaleString();

    // Log-Funktion zu Testzwecken
    console.log(time + ' - ' + user + ' - ' + this._checkXSS(message));
    let msgstream = {time: time, user: user, message: this._checkXSS(message)};
    let msgjson = JSON.stringify(msgstream);
    
    // Consolenausgabe:
    console.log(msgjson);
    // this.addToList(msgjson);

    // Auskommentiert, da nicht funktionierend
    //this.checkListCreation();

    // Ruft die Funktion auf, die die Nachricht versendet
    this._addToList(time, user, message);

    // Leert das Textfeld
    this.domElement.querySelector("textarea").value = '';

    // Lädt die Nachrichten auch hier noch einmal gesondert herunter
    await (this._renderListAsync());
  }

  /*  Funktionalität ist noch nicht sichergestellt.
      Zurzeit wird jedes Mal erneut der Versuch unternommen, die Liste zu erstellen.
      SharePoint liefert hier - zum Glück - HTTP500er-Meldungen, weshalb ich diese Funktion zunächst vernachlässigt habe.
      
      Bitte bei Zeiten unbedingt fertigstellen, um die Netzlast zu verringern!
  */ 
  public checkListCreation(list) {
    this._createList();
  }

  public _addToList(zeit, user, message) {
    var checkedMessage = this._checkXSS(message);
    pnp.sp.web.lists.getByTitle(this.properties.list).items.add({Time: zeit, User: user, Message: checkedMessage}).then(r=> console.log(r));
  }

  public _checkXSS(message) {
    var expr = /<(|\/|[^\/>][^>]+|\/[^>][^>]+)>/gi;
    var message_new = message.replace(expr, "");
    if (this.properties.allowhyperlinks) {
      var regex = /https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)/gi;
      var hyperlinks = [];
      hyperlinks = message.match(regex);
      if (typeof hyperlinks !== 'undefined' && hyperlinks != null) {
        hyperlinks.forEach(url => {
          console.log(message_new);
          message_new = message_new.replace(url, '<a href="'+url+'">'+url+'</a>');
        });
      }
      else {
        console.log(message_new);
      }
    }
    return message_new;
  }
  
  public _isUserModerator() {
    let result: boolean;
    let currentUser = this.context.pageContext.user;
    this.properties.moderators.forEach( mod => {
      console.log(mod.email+' - '+currentUser.email);
      if (String(mod.email) == String(currentUser.email)) {
        result = true;
      }
      else {
        result = false;
      }
    });
    return result;
  }

  /*  Auch hier wird (leider) noch bei jedem Seitenaufruf ein entsprechendes Feld in der Liste erstellt. 
      Daher wird diese Funktion nie aufgerufen. Alle Felder müssen - solange die Funktion nicht fertig ist - manuell in SharePoint hinzugefügt werden.
  */
  public async _setFieldTypes() {
    let isTimeThere = false;
    let isUserThere = false;
    let isMessageThere = false;
    let isAnswertoThere = false;
    let isPinnedThere = false;
    pnp.sp.web.lists.getByTitle(this.properties.list).fields.get().then(fields => {
      fields.forEach(field => {
        if (field.Title == "Time") {
          isTimeThere = true;
        }
        else if (field.Title == "User") {
          isUserThere = true;
        }
        else if (field.Title == "Message") {
          isMessageThere = true;
        }
        else if (field.Title == "AnswerTo") {
          isAnswertoThere = true;
        }
        else if (field.Title == "Pinned") {
          isPinnedThere = true;
        }
      });
    }).then(async fields => {
        if (!isTimeThere) {
          pnp.sp.web.lists.getByTitle(this.properties.list).fields.add("Time", "SP.FieldText", {"FieldTypeKind": 2, "Required": true}).then(async r => {
            if (!isUserThere) {
              pnp.sp.web.lists.getByTitle(this.properties.list).fields.add("User", "SP.FieldText", {"FieldTypeKind": 2, "Required": true}).then(async r2 => {
                if (!isMessageThere) {
                  pnp.sp.web.lists.getByTitle(this.properties.list).fields.add("Message", "SP.FieldText", {"FieldTypeKind": 2, "Required": true}).then(async r3 => {
                    if (!isAnswertoThere) {
                      pnp.sp.web.lists.getByTitle(this.properties.list).fields.add("AnswerTo", "SP.FieldText", {"FieldTypeKind": 2, "Required": false}).then(async r4 => {
                        if (!isPinnedThere) {
                          pnp.sp.web.lists.getByTitle(this.properties.list).fields.add("Pinned", "SP.FieldText", {"FieldTypeKind": 2, "Required": false});
                        }
                      });
                    }
                  });
                }
              });
            }
          });
        }
    });
  }

  // Erstellt die Liste über PnPjs
  // Der Wunsch, dass die Liste nur dann erstellt wird, wenn sie noch nicht existiert, ist auch hier noch nciht lauffähig. 
  public async _createList() {
    var listExists: boolean = false;
    pnp.sp.web.lists.get().then(lists => {
      lists.forEach(async list => {
        console.log(String(list.Title)+"--"+String(this.properties.list));
        if (String(list.Title) == String(this.properties.list)) {
          listExists = true;
        }
      });
    }).then(r => {console.log("listExists: "+listExists);}).then(async r => {
      if (listExists == false) {
        pnp.sp.web.lists.add(this.properties.list, 'basechat-list', 100, true).then(async r2 => {
          await this._setFieldTypes();
        }
        );
      }
    }
    );
  }
  
  // Lädt die Nachrichten in der Liste über einen REST-API-Call herunter.
  public _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('`+listName+`')/items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // Erzeugt für jede Nachricht das entsprechende HTML-Objekt und fügt es in den entsprechenden Bereich hinzu
  private _renderList(items: ISPList[]): void {

    messages = items.length;
    let html: string = '';
    let htmlPinned: string = '';
   
    items.forEach((item: ISPList) => {
      if (this._isUserModerator()) {
            // Nachrichten, die der Nutzer verschickt, werden messenger-typisch rechtsbündig und grün angezeigt
          if (item.User == this.context.pageContext.user.displayName) {
            if (!item.Pinned) {
              html += `
              <div class="${styles.bubble } ${styles.alt}" id="msg_${item.Id}">
                <div class="${styles.txt}">
                  <p class="${styles.message}" style="text-align: right;">${item.Message}</p><br>
                  <span class="${styles.timestamp}"><button class="delete ${styles.del_btn}" msg="${item.Id}">🗑</button> - <button class="pin ${styles.del_btn}" msg="${item.Id}">📌</button> - <button class="${styles.answer}" id="msg_${item.Id}">Antworten</button> ${item.Time}</span>
                </div>
              </div>`;
            }
            else {
              htmlPinned += `
              <div class="${styles.bubble } ${styles.alt} ${styles.pinned}" id="msg_${item.Id}">
                <div class="${styles.txt}">
                  <p class="${styles.message}" style="text-align: right;">${item.Message}</p><br>
                  <span class="${styles.timestamp}"><button class="delete ${styles.del_btn}" msg="${item.Id}">🗑</button> - <button class="pin ${styles.del_btn}" msg="${item.Id}">📌</button> - <button class="${styles.answer}" id="msg_${item.Id}">Antworten</button> ${item.Time}</span>
                </div>
              </div>`;
            }
          }
          else {
            // Andere Nachrichten linksbündig und mit weißem Hintergrund
            if (!item.Pinned) {
            html += `
            <div class="${styles.bubble }">
              <div class="${styles.txt}">
                <p class="${styles.name}">${item.User}</p>
                <p class="${styles.message}">${item.Message}</p><br>
                <span class="${styles.timestamp}"><button class="delete ${styles.del_btn}" msg="${item.Id}">🗑</button> - <button class="pin ${styles.del_btn}" msg="${item.Id}">📌</button> - ${item.Time}</span>
              </div>
            </div>`;
            }
            else {
              htmlPinned += `
              <div class="${styles.bubble } ${styles.pinned}">
                <div class="${styles.txt}">
                  <p class="${styles.name}">${item.User}</p>
                  <p class="${styles.message}">${item.Message}</p><br>
                  <span class="${styles.timestamp}"><button class="delete ${styles.del_btn}" msg="${item.Id}">🗑</button> - <button class="pin ${styles.del_btn}" msg="${item.Id}">📌</button> - ${item.Time}</span>
                </div>
              </div>`;
            }
          }
      }
      else {
        // Nachrichten, die der Nutzer verschickt, werden messenger-typisch rechtsbündig und grün angezeigt
        if (item.User == this.context.pageContext.user.displayName) {
          if (!item.Pinned) {
            html += `
            <div class="${styles.bubble } ${styles.alt}" id="msg_${item.Id}">
              <div class="${styles.txt}">
                <p class="${styles.message}" style="text-align: right;">${escape(item.Message)}</p><br>
                <span class="${styles.timestamp}"><button class="${styles.answer}" id="msg_${item.Id}">Antworten</button> ${item.Time}</span>
              </div>
            </div>`;
          }
          else {
            htmlPinned += `
            <div class="${styles.bubble } ${styles.alt} ${styles.pinned}" id="msg_${item.Id}">
              <div class="${styles.txt}">
                <p class="${styles.message}" style="text-align: right;">${escape(item.Message)}</p><br>
                <span class="${styles.timestamp}"><button class="${styles.answer}" id="msg_${item.Id}">Antworten</button> ${item.Time}</span>
              </div>
            </div>`;
          }
          
        }
        else {
          // Andere Nachrichten linksbündig und mit weißem Hintergrund
          if (!item.Pinned) {
            html += `
            <div class="${styles.bubble }">
              <div class="${styles.txt}">
                <p class="${styles.name}">${item.User}</p>
                <p class="${styles.message}">${item.Message}</p><br>
                <span class="${styles.timestamp}">${item.Time}</span>
              </div>
            </div>`;
          }
          else {
            htmlPinned += `
            <div class="${styles.bubble } ${styles.pinned}">
              <div class="${styles.txt}">
                <p class="${styles.name}">${item.User}</p>
                <p class="${styles.message}">${item.Message}</p><br>
                <span class="${styles.timestamp}">${item.Time}</span>
              </div>
            </div>`;
          }
          
        }
      }
    });
  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    const pinnedContainer: Element = this.domElement.querySelector('#spPinnedContainer');
    listContainer.innerHTML = html;
    pinnedContainer.innerHTML = htmlPinned;

    const delete_btns = document.querySelectorAll(".delete");
    const pin_btns = document.querySelectorAll(".pin");
    delete_btns.forEach(el => {
      el.addEventListener('click', event => {
        var target = <HTMLElement> event.target;
        pnp.sp.web.lists.getByTitle(this.properties.list).items.getById(Number(target.getAttribute('msg'))).delete();
      });
    });
    pin_btns.forEach(el => {
      el.addEventListener('click', event => {
        var target = <HTMLElement> event.target;
        pnp.sp.web.lists.getByTitle(this.properties.list).items.getAll().then(async r => {
          r.forEach(message => {
            if (Number(message.Id) == Number(target.getAttribute('msg'))) {
              pnp.sp.web.lists.getByTitle(this.properties.list).items.getById(Number(target.getAttribute('msg'))).update({Pinned: 'true'});
            }
            else {
              pnp.sp.web.lists.getByTitle(this.properties.list).items.getById(Number(message.Id)).update({Pinned: ''});
            }
          });
        });
        console.log("Message pinned.");
      });
    });

    // Scrollt im Nachrichtenbereich automatisch nach unten
    this.domElement.querySelector('#spListContainer').scrollTo(0,this.domElement.querySelector('#spListContainer').scrollHeight);
  }

  // Koordiniert das Herunterladen der Nachrichten, pseudo-asynchron durch die then-Funktion
  private _renderListAsync(): void {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }


  protected get dataVersion(): Version {
  return Version.parse('1.0');
}

  // EInstellungen für das Webpart auch hier wieder im PropertyPane
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
              PropertyPaneTextField('title', {
                label: strings.TitleFieldLabel
              }),
              PropertyPaneTextField('list', {
                label: strings.ListFieldLabel
              }),
              PropertyFieldPeoplePicker('moderators', {
                label: strings.ModeratorsFieldLabel,
                initialData: this.properties.moderators,
                allowDuplicate: false,
                principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                onPropertyChange: this.onPropertyPaneFieldChanged,
                context: this.context,
                properties: this.properties,
                onGetErrorMessage: null,
                deferredValidationTime: 0,
                key: 'peopleFieldId'
              }),
              PropertyPaneCheckbox('autofetch', {
                text: 'Neue Nachrichten automatisch laden',
                checked: true
              }),
              PropertyPaneCheckbox('allowhyperlinks', {
                text: 'Links in Nachrichten erlauben',
                checked: false
              }),
              PropertyPaneTextField('emojiset', {
                label: strings.EmojisetFieldLabel,
                value: '😅'
              }),
              // Abrufintervall in Millisekunden
              PropertyPaneTextField('interval', {
                label: strings.IntervalFieldLabel
              })
            ]
          },
          {
            groupName: strings.StyleGroupName,
            groupFields: [
              PropertyFieldColorPicker('background', {
                label: strings.BackgroundFieldLabel,
                selectedColor: this.properties.background,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                isHidden: false,
                alphaSliderHidden: false,
                style: PropertyFieldColorPickerStyle.Full,
                iconName: 'Precipitation',
                key: 'background'
              })
            ]
          }
        ]
      }
    ]
  };
}
}
