import { JupyterFrontEnd, JupyterFrontEndPlugin } from '@jupyterlab/application';
import { Widget } from '@lumino/widgets';
import { IStatusBar } from '@jupyterlab/statusbar';
import { Dialog, showDialog } from '@jupyterlab/apputils';
import { ISettingRegistry } from '@jupyterlab/settingregistry';
import { ServerConnection } from '@jupyterlab/services';
import DOMPurify from 'dompurify';
// ===================================================================================

DOMPurify.addHook('afterSanitizeAttributes', (node: any) => {
  // set all elements owning target to target=_blank
  if ('target' in node) {
    node.setAttribute('target', '_blank');
  }
  if (node.tagName === 'A' && node.getAttribute('href')) {
    node.setAttribute('rel', 'noopener noreferrer');
  }
});

// Use DOMPurify to sanitize HTML content
function sanitizeHtml(html: string): string {
  return DOMPurify.sanitize(html, {
    ALLOWED_TAGS: ['a', 'b', 'i', 'u', 'strong', 'em', 'br', 'p', 'ul', 'ol', 'li', 'code', 'pre'],
    ALLOWED_ATTR: ['href', 'target', 'rel'],
    ALLOW_DATA_ATTR: false,
    ADD_ATTR: ['target']
    // Ensure external links open in new tab
  });
}

// Keep escapeHtml for non-HTML content that should be displayed as plain text
function escapeHtml(html: string): string {
  const div = document.createElement('div');
  div.appendChild(document.createTextNode(html));
  return div.innerHTML;
}

function dlog(...args: any[]) {
  //  console.log(...args);
}

class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ValidationError';
  }
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
      throw new ValidationError('Bad message username.');
    }
    if (!(typeof timestamp === 'string' && /^\d\d\d\d-\d\d-\d\dT\d\d:\d\d(:\d\d(.\d\d\d(\d\d\d)?)?)?$/.test(timestamp))) {
      throw new ValidationError('Bad message timestamp.');
    }
    if (!(typeof expires === 'string' && /^\d+-\d\d:\d\d:\d\d$/.test(expires))) {
      throw new ValidationError('Bad expires time.');
    }
    if (!['debug', 'info', 'notice', 'warning', 'error', 'critical'].includes(level)) {
      throw new ValidationError('Bad level setting.');
    }
    if (!(typeof message === 'string' && message.length > 0 && message.length <= 1024)) {
      throw new ValidationError('Bad message content.');
    }
    // Escape metadata fields (these should be plain text)
    this.username = escapeHtml(username);
    this.timestamp = escapeHtml(timestamp);
    this.expires = escapeHtml(expires);
    this.level = escapeHtml(level);
    // Sanitize message content (this can contain safe HTML)
    this.message = sanitizeHtml(message);
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
      throw new ValidationError('Bad title.');
    }
    if (!Array.isArray(messages) || messages.length <= 0 || messages.length > 16) {
      // currently 5 on  server
      throw new ValidationError('Bad messages array.');
    }
    if (!messages.every(msg => msg instanceof Message)) {
      throw new ValidationError('Bad messages.');
    }
    // Sanitize title (allow basic formatting)
    this.title = sanitizeHtml(title);
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
      throw new ValidationError('Bad popup type.');
    }
    if (!Array.isArray(blocks)) {
      throw new ValidationError('Bad blocks array.');
    }
    if (!blocks.every(block => block instanceof MessageBlock)) {
      throw new ValidationError('Bad blocks.');
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
  // All parameters should be checked and sanitized to make them safe
  // for display in the browser.
  if (!Array.isArray(jsonData.blocks)) {
    throw new ValidationError('Bad blocks array.');
  }
  const blocks = jsonData.blocks.map((blockData: MessageBlock) => {
    const bdmessages = blockData.messages;
    if (!Array.isArray(bdmessages) || bdmessages.length <= 0 || bdmessages.length > 16) {
      // currently 5 on  server
      throw new ValidationError('Bad messages array.');
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
  // tracks retry backoff state
  retryDelay: number;
  maxRetryDelay: number;
  initialRetryDelay: number;
  // tracks service state for visual feedback
  serviceState: 'normal' | 'degraded' | 'failed';

  // takes the statusbar that we will add to as only parameter
  public constructor(statusbar: IStatusBar) {
    this.openAnnouncementButton = null;
    this.last_rendered = '';
    this.statusbar = statusbar;
    this.newAnnouncement = false;
    this.initialRetryDelay = 10000; // 10 seconds
    this.maxRetryDelay = 600000; // 10 minutes
    this.retryDelay = this.initialRetryDelay;
    this.serviceState = 'normal';
  }

  // Helper method to check if status code indicates permission issues
  private isPermissionError(status: number): boolean {
    return status === 401 || status === 403 || status === 407;
  }

  // Helper method to check if we should retry based on status code
  private shouldRetry(status: number): boolean {
    // Retry on permission errors and server errors (5xx)
    return this.isPermissionError(status) || (status >= 500 && status < 600);
  }

  // Update button text and styling based on service state
  private updateButtonForServiceState() {
    if (!this.openAnnouncementButton) {
      return;
    }
    switch (this.serviceState) {
      case 'normal':
        if (!this.newAnnouncement) {
          this.openAnnouncementButton.node.textContent = 'Announcements';
        } else {
          this.openAnnouncementButton.node.textContent = '⚠️ Click for Announcements';
        }
        this.openAnnouncementButton.node.style.color = '';
        break;

      case 'degraded':
        this.openAnnouncementButton.node.textContent = '🔄 Announcements (Retrying...)';
        this.openAnnouncementButton.node.style.color = 'orange';
        break;

      case 'failed':
        this.openAnnouncementButton.node.textContent = '❌ Announcements (Unavailable)';
        this.openAnnouncementButton.node.style.color = 'red';
        break;
    }
  }

  // Update button tooltip based on service state
  private updateButtonTooltip() {
    if (!this.openAnnouncementButton) {
      return;
    }
    switch (this.serviceState) {
      case 'normal':
        this.openAnnouncementButton.node.title = 'Click to view announcements';
        break;
      case 'degraded':
        this.openAnnouncementButton.node.title = `Announcement service experiencing issues. Retrying in ${this.retryDelay / 1000}s...`;
        break;
      case 'failed':
        this.openAnnouncementButton.node.title = 'Announcement service unavailable. Will retry periodically.';
        break;
    }
  }

  // Show recovery notification when service is restored
  private showRecoveryNotification() {
    if (!this.openAnnouncementButton) {
      return;
    }
    // Temporarily show recovery message
    this.openAnnouncementButton.node.textContent = '✅ Announcements Restored';
    this.openAnnouncementButton.node.style.color = 'green';
    this.openAnnouncementButton.node.title = 'Announcement service has been restored';

    // Revert after 3 seconds
    setTimeout(() => {
      this.updateButtonForServiceState();
      this.updateButtonTooltip();
    }, 3000);
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
          Authorization: 'token ' + serverSettings.token,
          'Content-Type': 'application/json',
          'Cache-Control': 'no-cache'
        },
        cache: 'no-cache'
      });
      dlog('Response: ', response);

      // Check if response is not ok and should be retried
      if (!response.ok && this.shouldRetry(response.status)) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      // If we get here, either response is ok or it's a non-retryable error
      if (!response.ok) {
        console.warn(`Non-retryable error fetching announcements: ${response.status} ${response.statusText}`);
        // Reset retry delay for successful connection (even if non-retryable error)
        this.retryDelay = this.initialRetryDelay;
        const wasServiceDegraded = this.serviceState !== 'normal';
        this.serviceState = 'normal';

        // Show recovery notification if we were previously degraded
        if (wasServiceDegraded) {
          this.showRecoveryNotification();
        }

        // Hide button on non-retryable errors
        if (this.openAnnouncementButton) {
          this.openAnnouncementButton.node.textContent = '';
        }
        setTimeout(() => {
          this.updateAnnouncements(url, n);
        }, n);
        return;
      }

      const jsonData = await response.json();
      dlog('jsonData:', jsonData);

      const announcements = jsonToAnnouncementsData(jsonData);
      const rendered = announcements.toHtml();
      dlog('Rendered:', rendered);

      // Success case - check if service was previously degraded
      const wasServiceDegraded = this.serviceState !== 'normal';
      this.serviceState = 'normal';
      this.retryDelay = this.initialRetryDelay;

      // check to see if the data is new
      if (rendered !== this.last_rendered) {
        this.newAnnouncement = true;
        this.last_rendered = rendered;
      }

      // if we have an announcement display a button to get the announcements
      if (this.last_rendered.length > 0) {
        this.createAnnouncementsButton(this.newAnnouncement);

        // Show recovery notification if we were previously degraded
        if (wasServiceDegraded) {
          this.showRecoveryNotification();
        } else {
          // Update button state normally
          this.updateButtonForServiceState();
          this.updateButtonTooltip();
        }

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
      // Handle validation errors separately from network errors
      if (e instanceof ValidationError) {
        console.warn(`A validation error occurred while processing announcements: ${e.message}`);
        console.warn('The service will continue to poll for new announcements at the normal interval.');

        // Schedule the next poll without engaging backoff logic
        setTimeout(() => {
          this.updateAnnouncements(url, n);
        }, n);
        return;
      }

      // there was an error with fetching
      console.warn(`Error fetching announcements: ${e}`);

      // Update service state based on retry attempts
      if (this.retryDelay >= this.maxRetryDelay) {
        console.error(`Giving up on fetching announcements after reaching max retry delay of ${this.maxRetryDelay}ms`);
        this.serviceState = 'failed';

        // Ensure button exists to show failed state
        if (!this.openAnnouncementButton) {
          this.createAnnouncementsButton(false);
        }
        this.updateButtonForServiceState();
        this.updateButtonTooltip();

        // Reset retry delay and continue with normal polling interval
        this.retryDelay = this.initialRetryDelay;
        setTimeout(() => {
          this.updateAnnouncements(url, n);
        }, n);
        return;
      }

      // Set degraded state and update button
      this.serviceState = 'degraded';

      // Ensure button exists to show degraded state
      if (!this.openAnnouncementButton) {
        this.createAnnouncementsButton(false);
      }
      this.updateButtonForServiceState();
      this.updateButtonTooltip();

      // Use exponential backoff for retry
      console.log(`Retrying in ${this.retryDelay}ms...`);
      setTimeout(() => {
        this.updateAnnouncements(url, n);
      }, this.retryDelay);

      // Double the retry delay for next time, but don't exceed max
      this.retryDelay = Math.min(this.retryDelay * 2, this.maxRetryDelay);
      return;
    }

    // wait n microseconds and check again (normal polling)
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
