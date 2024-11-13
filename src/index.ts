import { JupyterFrontEnd, JupyterFrontEndPlugin } from '@jupyterlab/application';
import { Widget } from '@lumino/widgets';
import { IStatusBar } from '@jupyterlab/statusbar';
import { Dialog, showDialog } from '@jupyterlab/apputils';
import { ISettingRegistry } from '@jupyterlab/settingregistry';

import { ServerConnection } from '@jupyterlab/services';

// ===================================================================================

// Use the browser's built-in functionality to quickly and safely escape the string
function escapeHtml(html: string): string {
  const div = document.createElement('div');
  div.appendChild(document.createTextNode(html));
  return div.innerText;
}

function dlog(...args: any[]) {
//  console.log(...args);
}

function log_and_throw(...args: any[]) {
  console.log(...args);
  throw new Error(...args);
}

// ===================================================================================

// type Levels = 'debug' | 'info' | 'notice' | 'warning' | 'error' | 'critical';

class Message {
  username: string;
  timestamp: string;
  expires: string;
  level: string;
  message: string;

  constructor(username: string, timestamp: string, expires: string, level: string, message: string) {
    if (!(typeof username === 'string' && username.length > 0 && username.length <= 32 && /^[a-zA-Z0-9_@.-]+$/.test(username))) {
      log_and_throw('Bad message username.');
    }
    if (!(typeof timestamp === 'string' && /^\d\d\d\d-\d\d-\d\dT\d\d:\d\d(:\d\d(.\d\d\d(\d\d\d)?)?)?$/.test(timestamp))) {
      log_and_throw('Bad message timestamp.');
    }
    if (!(typeof expires === 'string' && /^\d+\-\d\d:\d\d:\d\d$/.test(expires))) {
      log_and_throw('Bad expires time.');
    }
    if (!['debug', 'info', 'notice', 'warning', 'error', 'critical'].includes(level)) {
      log_and_throw('Bad level setting.');
    }
    if (!(typeof message === 'string' && message.length > 0 && message.length <= 1024)) {
      log_and_throw('Bad message content.');
    }
    this.username = escapeHtml(username);
    this.timestamp = escapeHtml(timestamp);
    this.expires = escapeHtml(expires);
    this.level = escapeHtml(level);
    this.message = escapeHtml(message);
  }

  fmtTimestamp(): string {
    return `<td class='announce-timestamp'>[${this.timestamp.split('.')[0]}]</td>`;
  }

  fmtLevel(): string {
    return `<td class='announce-${this.level}'>${this.level.toUpperCase()}</td>`;
  }

  fmtMessage(): string {
    return `<td class='announce-text'>${this.message}</td>`;
  }

  toHtml(): string {
    return `<tr>${this.fmtTimestamp()} ${this.fmtLevel()} <td> &nbsp; </td> ${this.fmtMessage()}</tr>`;
  }
}

class MessageBlock {
  title: string;
  messages: Message[];

  constructor(title: string, messages: Message[]) {
    if (typeof title !== 'string' || title.length <= 0 || title.length > 128) {
      log_and_throw('Bad title.');
    }
    if (!Array.isArray(messages) || messages.length <= 0 || messages.length > 16) {
      // currently 5 on  server
      log_and_throw('Bad messages array.');
    }
    if (!messages.every(msg => msg instanceof Message)) {
      log_and_throw('Bad messages.');
    }
    this.title = escapeHtml(title);
    this.messages = messages;
  }

  toHtml(): string {
    if (this.messages.length === 0) {
      return '';
    } else {
      return `<table class="announcement-block">
                        <theader>
                            <p class='announce-block-title'>${this.title}</p>
                        </theader>
                        <tbody>
                            ${this.messages.map(msg => msg.toHtml()).join('\n')}
                        </tbody>
                    </table>`;
    }
  }
}

class AnnouncementsData {
  popup: boolean;
  timestamp: string;
  blocks: MessageBlock[];

  constructor(popup: boolean, timestamp: string, blocks: MessageBlock[]) {
    if (typeof popup !== 'boolean') {
      log_and_throw('Bad popup type.');
    }
    if (!Array.isArray(blocks)) {
      log_and_throw('Bad blocks array.');
    }
    if (!blocks.every(block => block instanceof MessageBlock)) {
      log_and_throw('Bad blocks.');
    }
    this.popup = popup;
    this.timestamp = escapeHtml(timestamp);
    this.blocks = blocks;
  }

  toHtml(): string {
    if (this.blocks.length === 0) {
      return '';
    } else {
      return `<div class="announcement">
                ${this.blocks.map(block => block.toHtml()).join('\n')}
                        <p/>
            </div>`;
    }
  }
}

function jsonToAnnouncementsData(jsonData: any): AnnouncementsData {
    // Given unchecked `jsonData`,  create an AnnouncementsData instance
    // checking each field as used or as lower level objects use them.
    // All parameters should be checked and escaped to make them safe
    // for display in the browser.
    if (!Array.isArray(jsonData.blocks)) {
        log_and_throw('Bad blocks array.');
    }
    const blocks = jsonData.blocks.map((blockData: MessageBlock) => {
        const bdmessages = blockData.messages;
        if (!Array.isArray(bdmessages) || bdmessages.length <= 0 || bdmessages.length > 16) {
            // currently 5 on  server
            log_and_throw('Bad messages array.');
        }
        const messages = bdmessages.map((msg: Message) => {
            return new Message(msg.username, msg.timestamp, msg.expires, msg.level, msg.message);
        });
        return new MessageBlock(blockData.title, messages);
    });
  return new AnnouncementsData(jsonData.popup, jsonData.timestamp, blocks);
}

// ===================================================================================

// Class that handles all the announcements refresh information and methods
class RefreshAnnouncements {
  // tracks the button to show announcements
  openAnnouncementButton: Widget;
  // this tracks the current stored announcement
  last_rendered: string;
  // this is the status bar at the bottom of the screen
  statusbar: IStatusBar;
  // this tracks whether or not the user has seen the announcement
  // it determines whether or not to show the yellow alert emoji
  newAnnouncement: boolean;

  // takes the statusbar that we will add to as only parameter
  public constructor(statusbar: IStatusBar) {
    this.openAnnouncementButton = null;
    this.last_rendered = '';
    this.statusbar = statusbar;
    this.newAnnouncement = false;
  }

  // fetches the announcements data every n microseconds from the given url
  // creates and destroys the announcement button based on result of fetch
  async updateAnnouncements(url: string, n: number) {
    try {
      dlog(`Updating announcements from ${url} every ${n} microseconds`);
      const serverSettings = ServerConnection.makeSettings();
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': 'token ' + serverSettings.token,
          'Content-Type': 'application/json',
          'Cache-Control': 'no-cache'
        },
        cache: 'no-cache'
      });
      dlog('Response: ', response);

      const jsonData = await response.json();
      dlog('jsonData:', jsonData);

      const announcements = jsonToAnnouncementsData(jsonData);
      const rendered = announcements.toHtml();
      dlog('Rendered:', rendered);

      // check to see if the data is new
      if (rendered !== this.last_rendered) {
        this.newAnnouncement = true;
        this.last_rendered = rendered;
      }

      // if we have an announcement display a button to get the announcements
      if (this.last_rendered.length > 0) {
        this.createAnnouncementsButton(this.newAnnouncement);
        if (this.newAnnouncement && announcements.popup) {
          this.openAnnouncements();
        }
      } else {
        // otherwise hide any button present
        if (this.openAnnouncementButton) {
          this.openAnnouncementButton.node.textContent = '';
        }
      }
    } catch (e) {
      // there was an error with fetching
      if (this.openAnnouncementButton) {
        this.openAnnouncementButton.dispose();
      }
    }
    // wait n microseconds and check again
    setTimeout(() => {
      this.updateAnnouncements(url, n);
    }, n);
  }

  // creates/edits a button on the status bar to open the announcements modal
  createAnnouncementsButton(newAnnouncement: boolean) {
    // class used to create the open announcements button
    class ButtonWidget extends Widget {
      public constructor(
        announcementsObject: RefreshAnnouncements,
        newAnnouncement: boolean,
        options = { node: document.createElement('span') }
      ) {
        super(options);
        this.node.classList.add('open-announcements');

        // when the button is clicked:
        // mark the announcement as no longer new
        // open the announcement in a modal
        // edit the announcement button (to get rid of yellow warning emoji)
        this.node.onclick = () => {
          announcementsObject.newAnnouncement = false;
          announcementsObject.openAnnouncements();
          announcementsObject.createAnnouncementsButton(announcementsObject.newAnnouncement);
        };
      }
    }

    // if the open announcements button isn't on the status bar
    if (!this.openAnnouncementButton) {
      // creates the open annonucements button
      this.openAnnouncementButton = new ButtonWidget(this, this.newAnnouncement);
      // places the button on the status bar
      this.statusbar.registerStatusItem('new-announcement', {
        align: 'left',
        item: this.openAnnouncementButton
      });
    }

    // labels the button if the announcement based on if it is new or not
    if (!newAnnouncement) {
      this.openAnnouncementButton.node.textContent = 'Announcements';
    } else {
      this.openAnnouncementButton.node.textContent = '⚠️ Click for Announcements';
    }
  }

  // creates and open the modal with the announcement in it
  async openAnnouncements() {
    // because the user has click to read the announcement
    // it is no longer new to them
    this.newAnnouncement = false;

    // create the inner body of the announcement popup
    const body = document.createElement('span');
    body.innerHTML = this.last_rendered;
    body.classList.add('announcement');
    const widget = new Widget();
    widget.node.appendChild(body);

    // show the modal popup with the announcement
    void showDialog({
      title: 'Announcements',
      body: widget,
      buttons: [Dialog.okButton({ label: 'Close' })]
    });
  }
}

const PLUGIN_ID = 'stsci-announce:plugin';

/**
 * Initialization data for the stsci-announce extension.
 */
const extension: JupyterFrontEndPlugin<void> = {
  id: PLUGIN_ID,
  autoStart: true,
  requires: [IStatusBar, ISettingRegistry],
  activate: async (app: JupyterFrontEnd, statusBar: IStatusBar, settingRegistry: ISettingRegistry) => {
    console.log('JupyterLab extension stsci-announce is activated!');

    const settings = await settingRegistry.load(PLUGIN_ID);
    const apiUrl = settings.get('url').composite as string;
    const refreshInterval = settings.get('refresh-interval').composite as number;

    console.log(`Fetching announcements from ${apiUrl} every ${refreshInterval} milliseconds`);

    const myObject = new RefreshAnnouncements(statusBar);
    myObject.updateAnnouncements(apiUrl, refreshInterval);
  }
};

export default extension;
