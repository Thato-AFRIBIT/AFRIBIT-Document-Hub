/* eslint-disable require-atomic-updates, @typescript-eslint/no-explicit-any */
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import styles from './DocumentHubWebPart.module.scss';
import * as strings from 'DocumentHubWebPartStrings';
// import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { GraphService } from './GraphService';

export interface IDocumenthubWebPartProps {
  description: string;
}

export default class DocumenthubWebPart extends BaseClientSideWebPart<IDocumenthubWebPartProps> {


  private _siteId: string | undefined;
  private _driveId: string | undefined;

  // Breadcrumb state tracker for folder navigation
  private _breadcrumb: { id: string; name: string }[] = [];


  /**
   * Returns a skeleton grid for loading state.
   */
  private getLoaderHtml(): string {
    return `
      <div class="${styles.skeletonGrid}">
        ${Array(8).fill(`<div class="${styles.skeletonTile}"></div>`).join('')}
      </div>
    `;
  }

  /**
   * Throttles a function to run at most once every `wait` milliseconds.
   */
  private throttle(fn: (...args: unknown[]) => void, wait: number): (...args: unknown[]) => void {
    let last = 0;
    return (...args: unknown[]) => {
      const now = Date.now();
      if (now - last >= wait) {
        last = now;
        fn.apply(this, args as unknown as any);
      }
    };
  }

  /**
   * Executes a Graph fetch function with loading indicator and error handling.
   */
  private async withLoading(fetchFn: () => Promise<string>): Promise<string> {
    const grid = this.domElement.querySelector(`.${styles.documentGrid}`);
    if (grid) {
      grid.innerHTML = this.getLoaderHtml();
    }
    try {
      return await fetchFn();
    } catch (err) {
      console.error(err);
      return `<p>Error loading documents.</p>`;
    }
  }

  /**
   * Logs a user action in the properties panel timeline.
   */
  private logEvent(action: string): void {
    const panel = this.domElement.querySelector(`.${styles.propertiesPanel}`);
    if (panel) {
      const timeline = panel.querySelector(`.${styles.timeline}`) as HTMLElement;
      const entry = document.createElement('div');
      entry.className = styles.timelineItem;
      entry.innerHTML = `
        <div class="${styles.timelineDot}"></div>
        <div class="${styles.timelineContent}">
          <strong>You</strong> ${action} at ${new Date().toLocaleString()}
        </div>`;
      if (timeline) {
        timeline.appendChild(entry);
      } else {
        panel.innerHTML += `
          <p><strong>Activity:</strong></p>
          <div class="${styles.timeline}">${entry.outerHTML}</div>`;
      }
    }
  }

  private _graphClient: MSGraphClientV3 | undefined;
  private _graphService!: GraphService;

  /**
   * Retrieves the SharePoint site ID for AFRIBIT432.
   */
  private async getSiteId(): Promise<string> {
    return this._graphService.getSiteId();
  }

  /**
   * Retrieves the default drive ID for a given site.
   */
  private async getDriveId(siteId: string): Promise<string> {
    return this._graphService.getDriveId(siteId);
  }

  private readonly _searchRowLimit: number = 50;
  private _loadingMore: boolean = false;

  // Delta API state for “All Documents”
  private _allDocsNextLink: string | undefined;

  // Client-side sort & filter state
  private _sortOption: 'modifiedDesc' | 'modifiedAsc' | 'nameAsc' | 'nameDesc' = 'modifiedDesc';
  private _filterLast7Days: boolean = false;




  protected async onInit(): Promise<void> {
    this._graphClient = await this.context.msGraphClientFactory.getClient('3');
    this._graphService = new GraphService(this._graphClient);
    // Verify Graph connectivity by fetching signed-in user profile
    this._graphService.testGraphCall()
      .then(profile => {
        console.log('GraphService testGraphCall succeeded:', profile);
      })
      .catch(err => {
        console.error('GraphService testGraphCall failed:', err);
      });
    this._siteId = await this._graphService.getSiteId();
    this._driveId = await this._graphService.getDriveId(this._siteId);
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.documentHubWrapper}">
      <div class="${styles.headerSection}">
        <div class="${styles.headerContent}">
          <h2>Document Hub</h2>
          <div class="${styles.searchBox}">
            <input type="text" placeholder="Search documents..." />
          </div>
          <div class="${styles.uploadButtonWrapper}">
            <button id="uploadButton">Upload Document</button>
            <input type="file" id="fileInput" style="display:none" />
          </div>
        </div>
      </div>

      <div class="${styles.mainContent}">
        <div class="${styles.contentWrapper}">
          <button class="${styles.sidebarToggle}" aria-label="Toggle menu">☰</button>
          <aside class="${styles.sidebar}">
            <ul>
              <li class="${styles.activeSidebarItem}">All Documents</li>
              <li>My Folders</li>
              <li>Shared With Me</li>
              <li>Recently Accessed</li>
            </ul>
          </aside>

          <div class="${styles.centerContent}">
            <div class="${styles.tabs} ${styles.tabHeaders}">
              <button class="${styles.tabButton} ${styles.activeTabButton}" data-tab="recentTab">Recent</button>
            </div>
            <div class="${styles.sortControls}">
              <label for="sortSelect">Sort:</label>
              <select id="sortSelect">
                <option value="modifiedDesc">Modified ↓</option>
                <option value="modifiedAsc">Modified ↑</option>
                <option value="nameAsc">Name A→Z</option>
                <option value="nameDesc">Name Z→A</option>
              </select>
              <label><input type="checkbox" id="filter7Days" /> Last 7 days</label>
            </div>
            <div class="${styles.documentGrid}">
              <!-- Dynamic content will be rendered here -->
            </div>
          </div>

          <aside class="${styles.propertiesPanel}">
            <h3>Document Details</h3>
            <p>Select a document to view properties.</p>
          </aside>

        </div>
      </div>

      <footer class="${styles.footer}">
        <p>Powered by Microsoft 365 and Graph API | Document Hub v1.0</p>
      </footer>
    </div>
    `;
    // Mobile sidebar toggle behavior
    const toggleBtn = this.domElement.querySelector(`.${styles.sidebarToggle}`);
    const sidebarEl = this.domElement.querySelector(`.${styles.sidebar}`);
    if (toggleBtn && sidebarEl) {
      toggleBtn.addEventListener('click', () => {
        sidebarEl.classList.toggle(styles.open);
      });
    }

    // New tab layout and event handling
    const tabButtons = this.domElement.querySelectorAll(`.${styles.tabButton}`);
    const documentGrid = this.domElement.querySelector(`.${styles.documentGrid}`);
    tabButtons.forEach(button => {
      button.addEventListener('click', async () => {
        // Update active tab styling
        tabButtons.forEach(btn => btn.classList.remove(styles.activeTabButton));
        button.classList.add(styles.activeTabButton);

        // Show loader
        documentGrid!.innerHTML = this.getLoaderHtml();

        // Determine which tab and load content
        const tab = button.getAttribute('data-tab');
        let html = '';
        switch (tab) {
          case 'recentTab':
            html = await this.withLoading(() => this.renderRecentDocuments());
            break;
          case 'allTab':
            html = await this.withLoading(() => this.renderAllDocuments());
            break;
          case 'accessTab':
            html = await this.withLoading(() => this.renderSharedWithMe());
            break;
          case 'versionTab':
            html = '<p>Version history not available here.</p>';
            break;
          case 'securityTab':
            html = '<p>Open a document to view security details.</p>';
            break;
        }
        documentGrid!.innerHTML = html;
      });
    });

    const uploadButton = this.domElement.querySelector('#uploadButton') as HTMLButtonElement;
    const fileInput = this.domElement.querySelector('#fileInput') as HTMLInputElement;

    if (uploadButton && fileInput) {
      uploadButton.addEventListener('click', () => {
        fileInput.click();
      });

      fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) {
          return;
        }

        fileInput.value = ""; // Clear file selection immediately

        try {
          await this._graphService.uploadFileToDrive(`/root:/${file.name}`, file);

          alert('Upload successful!');
          const documentGrid = this.domElement.querySelector(`.${styles.documentGrid}`);
          if (documentGrid) {
            documentGrid.innerHTML = await this.renderRecentDocuments();
          }
        } catch (error) {
          console.error('Upload failed', error);
          alert('Upload failed. See console for details.');
        }
      });
    }

    if (documentGrid) {
      documentGrid.innerHTML = this.getLoaderHtml();
      this.withLoading(() => this.renderRecentDocuments())
        .then(html => {
          documentGrid!.innerHTML = html;
        })
        .catch(error => {
          console.error('Failed to load recent documents', error);
          documentGrid!.innerHTML = '<p>Error loading documents.</p>';
        });

      // Phase 2: Wire up the search input for debounced search
      const searchInput = this.domElement.querySelector(`.${styles.searchBox} input`) as HTMLInputElement;
      let searchDebounce: number;
      if (searchInput) {
        searchInput.addEventListener('input', () => {
          clearTimeout(searchDebounce);
          searchDebounce = window.setTimeout(async () => {
            const query = searchInput.value.trim();
            const gridEl = this.domElement.querySelector(`.${styles.documentGrid}`) as HTMLElement;
            if (!gridEl) return;
            gridEl.innerHTML = this.getLoaderHtml();
            const resultsHtml = query
              ? await this.withLoading(() => this.renderSearchResults(query))
              : await this.withLoading(() => this.renderRecentDocuments());
            gridEl.innerHTML = resultsHtml;
          }, 500);
        });
      }

      // Sort/filter controls event handlers
      const sortSelect = this.domElement.querySelector('#sortSelect') as HTMLSelectElement;
      const filterCheckbox = this.domElement.querySelector('#filter7Days') as HTMLInputElement;
      if (sortSelect) {
        sortSelect.value = this._sortOption;
        sortSelect.addEventListener('change', async () => {
          this._sortOption = sortSelect.value as any;
          const gridEl = this.domElement.querySelector(`.${styles.documentGrid}`) as HTMLElement;
          gridEl.innerHTML = this.getLoaderHtml();
          gridEl.innerHTML = await this.withLoading(() => this.renderAllDocuments());
        });
      }
      if (filterCheckbox) {
        filterCheckbox.checked = this._filterLast7Days;
        filterCheckbox.addEventListener('change', async () => {
          this._filterLast7Days = filterCheckbox.checked;
          const gridEl = this.domElement.querySelector(`.${styles.documentGrid}`) as HTMLElement;
          gridEl.innerHTML = this.getLoaderHtml();
          gridEl.innerHTML = await this.withLoading(() => this.renderAllDocuments());
        });
      }

      documentGrid.addEventListener('click', async (event: MouseEvent) => {
        const target = event.target as HTMLElement;
        const tile = target.closest(`.${styles.documentTile}`) as HTMLElement;
        if (!tile) return;
        // Highlight selected document tile
        this.domElement.querySelectorAll(`.${styles.documentTile}`)
          .forEach(t => t.classList.remove(styles.selectedTile));
        tile.classList.add(styles.selectedTile);

        // --- PATCH: Folder navigation with breadcrumb ---
        const isFolder = tile.querySelector(`.${styles.documentIcon}`)?.textContent === "📁";
        if (isFolder) {
          const itemId = tile.getAttribute('data-id');
          if (!itemId) return;

          const folderName = tile.querySelector(`.${styles.documentTitle}`)?.textContent || 'Folder';
          this._breadcrumb.push({ id: itemId, name: folderName });

          documentGrid!.innerHTML = this.getLoaderHtml();
          
          await this.loadFolderContents(itemId);
          return;
        }
        // --- END PATCH ---

        const webUrl = tile.getAttribute('data-weburl');
        if (target.closest(`.${styles.viewButton}`) && webUrl) {
          window.open(`${webUrl}?web=1`, '_blank');
          this.logEvent('viewed the document');
          return;
        }
        if (target.closest(`.${styles.editButton}`) && webUrl) {
          window.open(`${webUrl}?action=edit`, '_blank');
          this.logEvent('edited the document');
          return;
        }
        const itemId = tile.getAttribute('data-id');
        const driveId = tile.getAttribute('data-drive-id');
        if (itemId) {
          this.loadDocumentProperties(itemId, driveId || undefined)
            .catch(error => console.error('Failed to load document properties', error));
        }
      });

      const sidebarItems = this.domElement.querySelectorAll(`.${styles.sidebar} li`);
      sidebarItems.forEach(item => {
        item.addEventListener('click', async () => {
          // Highlight active sidebar item
          sidebarItems.forEach(si => si.classList.remove(styles.activeSidebarItem));
          item.classList.add(styles.activeSidebarItem);
          const text = item.textContent || '';
          const gridEl = this.domElement.querySelector(`.${styles.documentGrid}`) as HTMLElement;
          if (!gridEl) return;
          // eslint-disable-next-line require-atomic-updates
          gridEl.innerHTML = this.getLoaderHtml();
          let html = '';
          if (text.includes('All Documents')) {
            // Reset delta paging
            this._allDocsNextLink = undefined;
            this._loadingMore = false;
            html = await this.withLoading(() => this.renderAllDocuments());
          } else if (text.includes('My Folders')) {
            html = await this.withLoading(() => this.renderMyFolders());
          } else if (text.includes('Shared With Me')) {
            html = await this.withLoading(() => this.renderSharedWithMe());
          } else {
            html = await this.withLoading(() => this.renderRecentDocuments());
          }
          // eslint-disable-next-line require-atomic-updates
          gridEl.innerHTML = html;
          // Infinite scroll setup for "All Documents"
          if (text.includes('All Documents')) {
            const throttledScroll = this.throttle(async () => {
              const threshold = 100;
              if (!this._loadingMore && gridEl.scrollTop + gridEl.clientHeight >= gridEl.scrollHeight - threshold) {
                this._loadingMore = true;
                const moreHtml = await this.renderAllDocuments();
                gridEl.insertAdjacentHTML('beforeend', moreHtml);
                this._loadingMore = false;
              }
            }, 200);
            gridEl.addEventListener('scroll', throttledScroll);
          }
        });
      });
    }
  }

  /**
   * Loads and renders the user's recent documents using Microsoft Graph `/me/drive/recent`.
   */
  private async renderRecentDocuments(): Promise<string> {
    if (!this._graphClient) {
      return `<p>Unable to load documents.</p>`;
    }

    try {
      // Retrieve user's recent documents via Graph
      const recentResp: any = await this._graphClient!
        .api('/me/drive/recent')
        .get();
      const itemsAll: any[] = recentResp.value || [];

      // Filter out folders
      let items = itemsAll.filter((i: any) => !i.folder);

      // Sort by most recent modification
      items.sort((a, b) =>
        new Date(b.lastModifiedDateTime || '').getTime() -
        new Date(a.lastModifiedDateTime || '').getTime()
      );

      // Limit to top 10
      items = items.slice(0, 10);

      if (!items.length) {
        return `<p>No recent documents found.</p>`;
      }

      return items.map(item => `
        <div class="${styles.documentTile}" data-id="${item.id}" data-drive-id="${item.parentReference?.driveId}" data-weburl="${item.webUrl}">
          <div class="${styles.documentIcon}">${item.folder ? "📁" : "📄"}</div>
          <div class="${styles.documentTitle}">${item.name}</div>
          <div class="${styles.documentMeta}">Modified: ${new Date(item.lastModifiedDateTime || "").toLocaleString()}</div>
          <div class="${styles.documentMeta}">Owner: ${item.lastModifiedBy?.user?.displayName || "Unknown"}</div>
          <div class="${styles.tileActions}">
            <button class="${styles.viewButton}">View</button>
            <button class="${styles.editButton}">Edit</button>
          </div>
        </div>
      `).join('');
    } catch (error) {
      console.error('Error fetching recent documents', error);
      return `<p>Error loading documents.</p>`;
    }
  }


  /**
   * Renders search results from personal OneDrive.
   */
  private async renderSearchResults(query: string): Promise<string> {
    try {
      const items = await this._graphService.searchDrive(query);
      if (!items.length) {
        return `<p>No documents found for "${query}".</p>`;
      }
      return items.map(item => `
        <div class="${styles.documentTile}" data-id="${item.id}" data-drive-id="${item.parentReference?.driveId}" data-weburl="${item.webUrl}">
          <div class="${styles.documentIcon}">📄</div>
          <div class="${styles.documentTitle}">${item.name}</div>
          <div class="${styles.documentMeta}">Modified: ${new Date(item.lastModifiedDateTime || "").toLocaleString()}</div>
          <div class="${styles.documentMeta}">Owner: ${item.lastModifiedBy?.user?.displayName || "Unknown"}</div>
          <div class="${styles.tileActions}">
            <button class="${styles.viewButton}">View</button>
            <button class="${styles.editButton}">Edit</button>
          </div>
        </div>
      `).join('');
    } catch (error) {
      console.error('Error searching documents', error);
      return `<p>Error searching documents.</p>`;
    }
  }

  private async loadDocumentProperties(itemId: string, driveId?: string): Promise<void> {
    // Normalize driveId: treat empty or whitespace-only as undefined
    const resolvedDriveId = driveId?.trim() ? driveId : undefined;
    // Determine the drive ID to use for Graph calls
    const serviceDriveId = resolvedDriveId || await this.getDriveId(await this.getSiteId());
    if (!this._graphClient) return;

    try {
      // Batch request to fetch item metadata only (no audit logs)
      const batchRequests = [
        {
          id: "item",
          method: "GET",
          url: `/drives/${serviceDriveId}/items/${itemId}?$expand=listItem($expand=fields)`
        }
      ];
      const batchResponse: any = await this._graphClient!
        .api(`/$batch`)
        .post({ requests: batchRequests });

      // Parse batch responses
      const response = batchResponse.responses.find((r: any) => r.id === "item")?.body;
      // const auditLogs = batchResponse.responses.find((r: any) => r.id === "audit")?.body.value || [];

      // Check if the document exists
      if (!response) {
        alert('Document not found or unavailable.');
        return;
      }

      // Determine a friendly file type
      const rawType = response.file?.mimeType || '';
      const fileName = response.name || '';
      let friendlyType: string;
      if (response.folder) {
        friendlyType = 'Folder';
      } else {
        const ext = fileName.split('.').pop()?.toLowerCase() || '';
        switch (ext) {
          case 'doc': case 'docx':
            friendlyType = 'Word document';
            break;
          case 'xls': case 'xlsx':
            friendlyType = 'Excel spreadsheet';
            break;
          case 'ppt': case 'pptx':
            friendlyType = 'PowerPoint presentation';
            break;
          case 'pdf':
            friendlyType = 'PDF document';
            break;
          case 'txt':
            friendlyType = 'Text file';
            break;
          case 'jpg': case 'jpeg':
            friendlyType = 'JPEG image';
            break;
          case 'png':
            friendlyType = 'PNG image';
            break;
          default:
            friendlyType = ext ? ext.toUpperCase() + ' file' : rawType || 'Unknown';
        }
      }

      // Fetch the AFRIBIT432 SharePoint site by path
      const siteId = await this.getSiteId();

      const panel = this.domElement.querySelector(`.${styles.propertiesPanel}`);
      if (panel !== null) {  // Ensure panel is not null
        let metadataHtml = `
          <h3>${response.name || "Unknown"}</h3>
          <p><strong>Size:</strong> ${typeof response.size === 'number' ? (response.size / 1024).toFixed(2) : "0.00"} KB</p>
          <p><strong>Type:</strong> ${friendlyType}</p>
          <p><strong>Modified:</strong> ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : "Unknown"}</p>
          <p><strong>Owner:</strong> ${response.lastModifiedBy?.user?.displayName || "Unknown"}</p>
        `;

        // Insert metadata check immediately after metadataHtml definition
        const fields = response.listItem?.fields || {};
        const hasMetadata = Boolean(fields.Department || fields.Project || fields.Category);
        const driveType = response.parentReference?.driveType;
        // Only show Upload button for OneDrive (personal/business) items without metadata
        if (driveType && driveType !== 'documentLibrary' && !hasMetadata) {
          metadataHtml += `<button id="uploadToOneDriveButton">Upload to OneDrive</button>`;
        }

        // Build metadata dropdowns
        // Department options (static list you already created)
        const deptChoices = ['Finance','Human Resources','IT & Infrastructure','Marketing','Operations','Legal'];
        const deptOptions = deptChoices.map(d => `<option value="${d}" ${d === fields.Department ? 'selected' : ''}>${d}</option>`).join('');

        // Category options (Internal/Public/Confidential/Restricted)
        const catChoices = ['Internal','Public','Confidential','Restricted'];
        const catOptions = catChoices.map(c => `<option value="${c}" ${c === fields.Category ? 'selected' : ''}>${c}</option>`).join('');

        // Define the shape of project items
        interface ProjectListItem {
          fields: {
            Title: string;
          };
        }
        // Fetch Projects from your Projects list on that site
        const projectsResponse = await this._graphClient!
          .api(`/sites/${siteId}/lists/Projects/items?$expand=fields`)
          .get() as { value: ProjectListItem[] };
        const projects = projectsResponse.value || [];
        const projectOptions = projects
          .map((p: ProjectListItem) => {
            const name = p.fields.Title;
            return `<option value="${name}" ${name === fields.Project ? 'selected' : ''}>${name}</option>`;
          })
          .join('');

        // Prepare the classification controls HTML for Details tab
        const classificationControlsHtml = `
          <div class="metadataControls">
            <br>
            <h4>Classify Document</h4>
            <label for="deptPicker">Department:</label>
            <select id="deptPicker">${deptOptions}</select>
            <label for="projectPicker">Project:</label>
            <select id="projectPicker">${projectOptions}</select>
            <label for="catPicker">Category:</label>
            <select id="catPicker">${catOptions}</select>
            <button id="saveMetaButton">Save Metadata</button>
          </div>
        `;


        // Render properties panel as HTML tabs (DLP and Encryption Key sections removed)
        panel.innerHTML = `
          <div class="${styles.tabHeaders}">
            <button class="${styles.tabButton} ${styles.activeTabButton}" data-tab="detailsTab">Details</button>
            <button class="${styles.tabButton}" data-tab="accessTab">Access</button>
            <button class="${styles.tabButton}" data-tab="versionTab">Version</button>
            <button class="${styles.tabButton}" data-tab="securityTab">Security</button>
          </div>
          <div id="detailsTab" class="${styles.tabContent}">
            ${metadataHtml}
            ${classificationControlsHtml}
          </div>
          <div id="accessTab" class="${styles.tabContent}" style="display:none;">
            <div class="accessLogsContainer">
              <p><strong>Access:</strong></p>
              <div class="${styles.timeline}">
                <!-- lazy-loaded on tab click -->
              </div>
            </div>
          </div>
          <div id="versionTab" class="${styles.tabContent}" style="display:none;">
            <div class="versionHistoryContainer">
              <p><strong>Version History:</strong></p>
              <div class="${styles.versionHistoryWrapper}">
                <ul id="versionHistoryList">
                  <!-- versions lazy-loaded on tab click -->
                </ul>
              </div>
            </div>
          </div>
          <div id="securityTab" class="${styles.tabContent}" style="display:none;">
            <div class="securitySection">
              <p><strong>Sensitivity Label:</strong></p>
              <ul id="labelList"><li>None</li></ul>
            </div>
            <div class="securitySection">
              <p><strong>Retention Policy:</strong></p>
              <ul id="retentionList"><li>None</li></ul>
            </div>
            <div class="securitySection">
              <p><strong>Conditional Access:</strong></p>
              <ul id="caPolicyList"><li>None</li></ul>
            </div>
            <div class="securitySection">
              <p><strong>Information Barriers:</strong></p>
              <ul id="ibList"><li>None</li></ul>
            </div>
          </div>
        `;

        // Tab switching handlers
        const tabButtons = panel.querySelectorAll(`.${styles.tabButton}`);
        tabButtons.forEach(btn => {
          btn.addEventListener('click', () => {
            // Deactivate all buttons and hide contents
            tabButtons.forEach(b => {
              b.classList.remove(styles.activeTabButton);
              const contentId = b.getAttribute('data-tab');
              panel.querySelector(`#${contentId}`)!.setAttribute('style', 'display:none;');
            });
            // Activate clicked button and show its content
            btn.classList.add(styles.activeTabButton);
            const activeContent = btn.getAttribute('data-tab')!;
            panel.querySelector(`#${activeContent}`)!.setAttribute('style', 'display:block;');
          });
        });

        // Security tab handler
        const securityBtn = panel.querySelector(`button[data-tab="securityTab"]`);
        securityBtn?.addEventListener('click', async () => {
          // Deactivate all tabs and hide contents
          tabButtons.forEach(b => {
            b.classList.remove(styles.activeTabButton);
            const contentId = b.getAttribute('data-tab')!;
            panel.querySelector(`#${contentId}`)!.setAttribute('style', 'display:none;');
          });
          // Activate Security
          securityBtn.classList.add(styles.activeTabButton);
          const secPane = panel.querySelector('#securityTab') as HTMLElement;
          secPane.setAttribute('style', 'display:block;');

          // Determine the effective drive ID for security calls
          const effectiveDriveId = driveId?.trim() ? driveId : serviceDriveId;

          // --- PATCH: Sensitivity label rendering and edit workflow ---
          // Load and display the document's current sensitivity label
          const labelSection = panel.querySelector('#labelList') as HTMLElement;
          labelSection.innerHTML = ''; // clear existing
          // show current label as a bullet
          const current = await this._graphService.getSensitivityLabel(effectiveDriveId, itemId);
          const statusText = current.name || 'None';
          labelSection.innerHTML = `<li>${statusText}</li>`;

          // If user is a SharePoint admin, render Edit button above status
          const isAdmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
          if (isAdmin) {
            const editBtn = document.createElement('button');
            editBtn.id = 'editLabelButton';
            editBtn.textContent = 'Edit Label';
            // insert button just above the label list
            labelSection.parentElement!.insertBefore(editBtn, labelSection);

            // create hidden picker and save button
            const picker = document.createElement('select');
            picker.id = 'labelPicker';
            picker.style.display = 'none';
            const saveBtn = document.createElement('button');
            saveBtn.id = 'saveLabelButton';
            saveBtn.textContent = 'Save Label';
            saveBtn.style.display = 'none';
            labelSection.parentElement!.appendChild(picker);
            labelSection.parentElement!.appendChild(saveBtn);

            editBtn.addEventListener('click', async () => {
              // fetch and populate labels
              const labels = await this._graphService.listSensitivityLabels();
              picker.innerHTML = labels.map(l => `<option value="${l.id}" ${l.name === current.name ? 'selected' : ''}>${l.name}</option>`).join('');
              picker.style.display = 'block';
              saveBtn.style.display = 'inline-block';
            });
            saveBtn.addEventListener('click', async () => {
              const selectedId = (picker as HTMLSelectElement).value;
              const selectedName = picker.selectedOptions[0].text;
              await this._graphService.assignSensitivityLabel(effectiveDriveId, itemId, selectedId);
              // update UI
              labelSection.innerHTML = `<li>${selectedName}</li>`;
              picker.style.display = 'none';
              saveBtn.style.display = 'none';
              this.logEvent(`assigned sensitivity label: ${selectedName}`);
            });
          }
          // --- END PATCH ---

          // Load and render retention policy as a list
          const retentionList = panel.querySelector('#retentionList') as HTMLElement;
          try {
            const currentRetention = await this._graphService.getRetentionLabel(effectiveDriveId, itemId);
            retentionList.innerHTML = currentRetention?.name
              ? `<li>${currentRetention.name}</li>`
              : `<li>None</li>`;
          } catch (e) {
            console.error('Failed to load retention policy:', e);
            retentionList.innerHTML = '<li>Unavailable</li>';
          }

          // Conditional Access policies (live)
          const caList = panel.querySelector('#caPolicyList') as HTMLElement;
          try {
            const caPolicies = await this._graphService.listConditionalAccessPolicies();
            const count = caPolicies.length;
            caList.innerHTML = `<li>${count} conditional access polic${count === 1 ? 'y' : 'ies'} found</li>`;
          } catch (e) {
            console.error('Failed to load CA policies:', e);
            caList.innerHTML = '<li>Unavailable</li>';
          }

          // Information Barriers
          const ibList = panel.querySelector('#ibList') as HTMLElement;
          try {
            const barriers = await this._graphService.listInformationBarriers();
            if (barriers.length) {
              ibList.innerHTML = barriers.map(b => `<li>${b.displayName || b.id}</li>`).join('');
            } else {
              ibList.innerHTML = '<li>None</li>';
            }
          } catch (e) {
            console.error('Failed to load information barriers:', e);
            ibList.innerHTML = '<li>Unavailable</li>';
          }
        });

        // Accessed By tab lazy-load handler
        const accessBtn = panel.querySelector(`button[data-tab="accessTab"]`);
        accessBtn?.addEventListener('click', async () => {
          const container = panel.querySelector(`.${styles.timeline}`) as HTMLElement;
          container.innerHTML = '<p>Loading access records…</p>';
          try {
            // Fetch recent activities for the file (reduced to 5 for performance)
            const activitiesResp: any = await this._graphClient!
              .api(`/drives/${serviceDriveId}/items/${itemId}/activities?$top=5`)
              .get();
            // const activities = activitiesResp.value as any[];
            const activities = activitiesResp.value as any[];
            container.innerHTML = activities.length > 0
              ? activities.map(act => {
                  const actor = act.actor?.user?.displayName || 'Unknown';
                  let actionDesc = 'accessed';
                  if (act.action?.edit) {
                    actionDesc = 'edited';
                  } else if (act.action?.checkin) {
                    actionDesc = 'checked in';
                  } else if (act.action?.version?.newVersion) {
                    actionDesc = `versioned to v${act.action.version.newVersion}`;
                  }
                  const timeStr = act.times?.recordedDateTime
                    ? new Date(act.times.recordedDateTime).toLocaleString()
                    : 'Unknown time';
                  return `
                    <div class="${styles.timelineItem}">
                      <div class="${styles.timelineDot}"></div>
                      <div class="${styles.timelineContent}">
                        <strong>${actor}</strong> ${actionDesc} at ${timeStr}
                      </div>
                    </div>`;
                }).join('')
              : '<em>No access records found.</em>';
          } catch {
            container.innerHTML = '<em>Failed to load access records.</em>';
          }
        });

        // Version tab lazy-load handler
        const versionBtn = panel.querySelector(`button[data-tab="versionTab"]`);
        versionBtn?.addEventListener('click', async () => {
          const listEl = panel.querySelector('#versionHistoryList') as HTMLElement;
          listEl.innerHTML = '<li>Loading versions…</li>';
          try {
            const versionsResp: any = await this._graphClient!
              .api(`/drives/${serviceDriveId}/items/${itemId}/versions?$top=5`)
              .get();
            const versions = versionsResp.value as any[];
            listEl.innerHTML = versions.length > 0
              ? versions.map(v => `
                  <li>
                    Version ${v.id} - ${new Date(v.lastModifiedDateTime).toLocaleString()} by ${v.lastModifiedBy?.user?.displayName || 'Unknown'}
                  </li>`).join('')
              : '<li><em>No versions found.</em></li>';
          } catch {
            listEl.innerHTML = '<li><em>Failed to load versions.</em></li>';
          }
        });

        // --- CLASSIFICATION CONTROLS LOGIC (same as before) ---
        // Initialize Save/Edit mode based on existing metadata
        const saveButtonInit = panel.querySelector('#saveMetaButton') as HTMLButtonElement | null;
        const selectElements = Array.from(panel.querySelectorAll('.metadataControls select')) as HTMLSelectElement[];
        if (hasMetadata) {
          // start in read-only mode
          selectElements.forEach(el => { el.disabled = true; });
          if (saveButtonInit) {
            saveButtonInit.textContent = 'Edit Metadata';
          }
        } else {
          // start in edit mode
          if (saveButtonInit) {
            saveButtonInit.textContent = 'Save Metadata';
          }
        }

        setTimeout(() => {
          const saveButton = panel!.querySelector('.metadataControls #saveMetaButton') as HTMLButtonElement;
          if (!saveButton) {
            console.error('[DEBUG] Save button not found in DOM.');
            return;
          }
          // Helper to find or create a status element
          function getOrCreateStatusEl(): HTMLElement {
            let statusEl = panel!.querySelector('#saveStatus') as HTMLElement;
            if (!statusEl) {
              statusEl = document.createElement('div');
              statusEl.id = 'saveStatus';
              statusEl.style.marginTop = '8px';
              panel!.querySelector('.metadataControls')!.appendChild(statusEl);
            }
            return statusEl;
          }
          saveButton.addEventListener('click', async () => {
            // If in Edit mode, switch to save mode
            if (saveButton.textContent === 'Edit Metadata') {
              panel!.querySelectorAll('.metadataControls select').forEach((el: Element) => {
                (el as HTMLSelectElement).disabled = false;
              });
              saveButton.textContent = 'Save Metadata';
              return;
            }
            // Save Metadata mode
            saveButton.disabled = true;
            saveButton.textContent = 'Saving…';
            const statusEl = getOrCreateStatusEl();
            statusEl.textContent = 'Saving metadata…';
            try {
              const dept    = (panel!.querySelector('#deptPicker')   as HTMLSelectElement).value;
              const project = (panel!.querySelector('#projectPicker') as HTMLSelectElement).value;
              const category= (panel!.querySelector('#catPicker')     as HTMLSelectElement).value;
              const fieldsResponse = response.listItem?.fields || {};
              const updatedFields: Record<string,string> = {};
              if ('Department' in fieldsResponse) updatedFields.Department = dept;
              if ('Project'    in fieldsResponse) updatedFields.Project    = project;
              if ('Category'   in fieldsResponse) updatedFields.Category   = category;
              if (Object.keys(updatedFields).length === 0) {
                statusEl.textContent = 'No metadata fields available to update.';
                return;
              }
              await this._graphClient!
                .api(`/drives/${serviceDriveId}/items/${itemId}/listItem/fields`)
                .header('Content-Type', 'application/json')
                .patch(updatedFields);
              this.logEvent('updated DriveItem metadata');
              statusEl.textContent = 'Metadata saved successfully!';
              // Switch to read-only view
              panel!.querySelectorAll('.metadataControls select').forEach((el: Element) => {
                (el as HTMLSelectElement).disabled = true;
              });
              saveButton.textContent = 'Edit Metadata';
            } catch (err) {
              console.error('Metadata PATCH failed', err);
              statusEl.textContent = 'Metadata update failed.';
            } finally {
              saveButton.disabled = false;
              setTimeout(() => { statusEl.textContent = ''; }, 3000);
            }
          });
        }, 0);

        // Upload to OneDrive button logic
        const uploadButton = panel.querySelector('#uploadToOneDriveButton') as HTMLButtonElement;
        if (uploadButton) {
          uploadButton.addEventListener('click', async () => {
            if (!this._graphClient) return;
            try {
              await this._graphClient.api('/me/drive/root/children').post({
                name: response.name,
                file: {},
                "@microsoft.graph.conflictBehavior": "rename"
              });
              alert('File successfully uploaded to OneDrive to start tracking.');
            } catch (uploadError) {
              console.error('Upload to OneDrive failed', uploadError);
              alert('Upload to OneDrive failed. See console for details.');
            }
          });
        }

        // Version history events
        setTimeout(() => {
          // Hook up toggle button to expand/collapse
          const toggleBtnEl = panel?.querySelector('#toggleVersionsButton');
          const listElEl = panel?.querySelector('#versionHistoryList');
          if (toggleBtnEl instanceof HTMLButtonElement && listElEl instanceof HTMLElement) {
            toggleBtnEl.addEventListener('click', () => {
              const isExpanded = !listElEl.classList.toggle(styles.collapsed);
              toggleBtnEl.textContent = isExpanded ? 'Show less' : 'Show more';
            });
          }
          // Hook up restore buttons
          panel?.querySelectorAll(`.${styles.restoreButton}`).forEach((btn: HTMLElement) => {
            btn.addEventListener('click', async () => {
              const versionId = btn.getAttribute('data-version-id');
              if (versionId) {
                const restorePath = `/drives/${serviceDriveId}/items/${itemId}`;
                try {
                  await this._graphClient!.api(`${restorePath}/versions/${versionId}/restoreVersion`).post({});
                  alert(`Version ${versionId} restored.`);
                  this.loadDocumentProperties(itemId, driveId).catch(console.error);
                } catch (restoreError) {
                  console.error('Restore failed', restoreError);
                  alert('Restore failed. See console.');
                }
              }
            });
          });
        }, 0);
      }
    } catch (error) {
      console.error('Error loading document properties', error);
      const panel = this.domElement.querySelector(`.${styles.propertiesPanel}`);
      if (!panel) return;

      // In catch(error) block before using actualDriveId:
      // Attempt basic metadata fetch for fileName
      let fileName = 'document';
      try {
        const meta = await this._graphClient!.api(
          `/drives/${serviceDriveId}/items/${itemId}`
        ).get();
        fileName = meta.name || fileName;
      } catch {
        // ignore
      }

      // Build UI: always offer copy to OneDrive for SharePoint items
      panel.innerHTML = `
        <h3>${fileName}</h3>
        <p>This file is shared and not in your OneDrive. You can copy it to your drive:</p>
        <button id="copyToOneDriveButton">📤 Copy to OneDrive</button>
      `;
      panel.querySelector('#copyToOneDriveButton')?.addEventListener('click', async () => {
        try {
          await this._graphClient!.api(`/drives/${serviceDriveId}/items/${itemId}/copy`).post({
            parentReference: { id: 'root' },
            name: fileName
          });
          alert('Copy initiated. Check your OneDrive root.');
        } catch (copyError) {
          console.error('Copy failed', copyError);
          alert('Copy failed. See console for details.');
        }
      });
    }
  }

  /**
   * Renders “All Documents” using Graph Delta API and paging, with client-side sort/filter.
   */
  private async renderAllDocuments(): Promise<string> {
    if (!this._graphClient) {
      return `<p>Unable to load documents.</p>`;
    }
    try {
      const driveId = this._driveId!;
      // Build delta query or continue via nextLink
      const pageSize = this._searchRowLimit;
      const deltaUrl = this._allDocsNextLink
        ? this._allDocsNextLink
        : `/drives/${driveId}/root/delta?$top=${pageSize}`;
      // Fetch page of changes
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const resp: any = await this._graphClient.api(deltaUrl).get();
      let items = resp.value as any[];
      // --- Client-side filter ---
      if (this._filterLast7Days) {
        const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
        items = items.filter(i => {
          const m = i.lastModifiedDateTime ? new Date(i.lastModifiedDateTime).getTime() : 0;
          return m >= cutoff;
        });
      }
      // --- Client-side sort ---
      switch (this._sortOption) {
        case 'modifiedDesc':
          items.sort((a, b) =>
            new Date(b.lastModifiedDateTime || "").getTime() - new Date(a.lastModifiedDateTime || "").getTime()
          );
          break;
        case 'modifiedAsc':
          items.sort((a, b) =>
            new Date(a.lastModifiedDateTime || "").getTime() - new Date(b.lastModifiedDateTime || "").getTime()
          );
          break;
        case 'nameAsc':
          items.sort((a, b) => (a.name || "").localeCompare(b.name || ""));
          break;
        case 'nameDesc':
          items.sort((a, b) => (b.name || "").localeCompare(a.name || ""));
          break;
      }
      // Update nextLink for subsequent pages
      this._allDocsNextLink = resp['@odata.nextLink'];
      if (!items.length) {
        return `<p>No more documents.</p>`;
      }
      // Render items
      return items.map(item => `
        <div class="${styles.documentTile}" data-id="${item.id}" data-drive-id="${driveId}" data-weburl="${item.webUrl}">
          <div class="${styles.documentIcon}">${item.folder ? "📁" : "📄"}</div>
          <div class="${styles.documentTitle}">${item.name}</div>
          <div class="${styles.documentMeta}">Modified: ${item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : ''}</div>
          <div class="${styles.tileActions}">
            <button class="${styles.viewButton}">View</button>
            <button class="${styles.editButton}">Edit</button>
          </div>
        </div>
      `).join('');
    } catch (error) {
      console.error('Error fetching all documents via Delta API', error);
      return `<p>Error loading documents.</p>`;
    }
  }

  private async renderMyFolders(): Promise<string> {
    if (!this._graphClient) {
      return `<p>Unable to load folders.</p>`;
    }
    try {
      // Fetch the AFRIBIT432 site and drive IDs via service
      const siteId = await this.getSiteId();
      const driveId = await this.getDriveId(siteId);
      // List only folders in the Shared Documents root
      const items = await this._graphService.getFolders(driveId, 'General');
      if (!items.length) {
        return `<p>No folders found.</p>`;
      }
      items.sort((a, b) =>
        new Date(b.lastModifiedDateTime || "").getTime() - new Date(a.lastModifiedDateTime || "").getTime()
      );
      return items.map((item: any) => {
        if (!item || !item.id || !item.name) {
          return '';
        }
        return `
          <div class="${styles.documentTile}" data-id="${item.id}" data-weburl="${item.webUrl}">
            <div class="${styles.documentIcon}">📁</div>
            <div class="${styles.documentTitle}">${item.name}</div>
            <div class="${styles.documentMeta}">Modified: ${new Date(item.lastModifiedDateTime || "").toLocaleString()}</div>
            <div class="${styles.documentMeta}">Owner: ${item.lastModifiedBy?.user?.displayName || "Unknown"}</div>
            <div class="${styles.tileActions}">
              <button class="${styles.viewButton}">View</button>
              <button class="${styles.editButton}">Edit</button>
            </div>
          </div>
        `;
      }).join('');
    } catch (error) {
      console.error('Error fetching folders', error);
      return `<p>Error loading folders.</p>`;
    }
  }

  private async renderSharedWithMe(): Promise<string> {
    if (!this._graphClient) {
      return `<p>Unable to load shared documents.</p>`;
    }

    try {
      const items = await this._graphService.getSharedWithMe();

      if (!items.length) {
        return `<p>No shared documents found.</p>`;
      }

      return items.map((shared: any) => {
        const remote = shared.remoteItem as any;
        return `
          <div class="${styles.documentTile}" data-id="${remote.id}" ${remote.parentReference?.driveId ? `data-drive-id="${remote.parentReference.driveId}"` : ''} data-weburl="${remote.webUrl}">
            <div class="${styles.documentIcon}">🔗</div>
            <div class="${styles.documentTitle}">${remote.name}</div>
            <div class="${styles.documentMeta}">Owner: ${remote.createdBy?.user?.displayName || remote.lastModifiedBy?.user?.displayName || "Unknown"}</div>
            <div class="${styles.tileActions}">
              <button class="${styles.viewButton}">View</button>
              <button class="${styles.editButton}" disabled>Edit</button>
            </div>
          </div>
        `;
      }).join('');
    } catch (error) {
      console.error('Error fetching shared documents', error);
      return `<p>Error loading shared documents.</p>`;
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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

  /**
   * Loads the folder contents for a given itemId, renders tiles, and attaches breadcrumb and back button click handlers.
   */
  private async loadFolderContents(itemId: string): Promise<void> {
    const documentGrid = this.domElement.querySelector(`.${styles.documentGrid}`);
    if (!documentGrid) return;

    try {
      const siteId = await this.getSiteId();
      const driveId = await this.getDriveId(siteId);
      const response = await this._graphClient!.api(`/drives/${driveId}/items/${itemId}/children`).get();
      const folderItems = response.value as any[];
      // Sort folder items by most recently modified first
      folderItems.sort((a, b) =>
        new Date(b.lastModifiedDateTime || '').getTime() -
        new Date(a.lastModifiedDateTime || '').getTime()
      );

      const breadcrumbHtml = this._breadcrumb.map((crumb, index) => `
        <span class="${styles.breadcrumbItem}" data-index="${index}" style="cursor: pointer;">${crumb.name}</span>
      `).join('');

      const folderHtml = folderItems.length > 0
        ? folderItems.map(item => `
            <div class="${styles.documentTile}" data-id="${item.id}" data-drive-id="${driveId}" data-weburl="${item.webUrl}">
              <div class="${styles.documentIcon}">${item.folder ? "📁" : "📄"}</div>
              <div class="${styles.documentTitle}">${item.name}</div>
              <div class="${styles.documentMeta}">Modified: ${new Date(item.lastModifiedDateTime || "").toLocaleString()}</div>
              <div class="${styles.documentMeta}">Owner: ${item.lastModifiedBy?.user?.displayName || "Unknown"}</div>
              <div class="${styles.tileActions}">
                <button class="${styles.viewButton}">View</button>
                <button class="${styles.editButton}">Edit</button>
              </div>
            </div>
          `).join('')
        : "<p>No items found in this folder.</p>";

      // eslint-disable-next-line require-atomic-updates
      documentGrid.innerHTML = `
        <div class="${styles.breadcrumbWrapper}" style="margin-bottom: 10px;">
          <button class="${styles.backButton}">⬅️ Back</button> ${breadcrumbHtml}
        </div>
        ${folderHtml}
      `;

      // Attach breadcrumb click events
      this.domElement.querySelectorAll(`.${styles.breadcrumbItem}`).forEach(el => {
        el.addEventListener('click', async (e) => {
          const target = e.currentTarget as HTMLElement;
          const index = parseInt(target.getAttribute('data-index') || '0', 10);
          this._breadcrumb = this._breadcrumb.slice(0, index + 1);
          const crumb = this._breadcrumb[this._breadcrumb.length - 1];
          // eslint-disable-next-line require-atomic-updates
          documentGrid.innerHTML = this.getLoaderHtml();
          if (crumb) {
            await this.loadFolderContents(crumb.id);
          } else {
            this._breadcrumb = [];
            // eslint-disable-next-line require-atomic-updates
            documentGrid.innerHTML = await this.withLoading(() => this.renderRecentDocuments());
          }
        });
      });

      // Attach back button
      const backButton = this.domElement.querySelector(`.${styles.backButton}`) as HTMLButtonElement;
      if (backButton) {
        backButton.addEventListener('click', async () => {
          this._breadcrumb.pop();
          const last = this._breadcrumb[this._breadcrumb.length - 1];
          // eslint-disable-next-line require-atomic-updates
          documentGrid.innerHTML = this.getLoaderHtml();
          if (last) {
            await this.loadFolderContents(last.id);
          } else {
            this._breadcrumb = [];
            // eslint-disable-next-line require-atomic-updates
            documentGrid.innerHTML = await this.withLoading(() => this.renderRecentDocuments());
          }
        });
      }

    } catch (error) {
      console.error('Error loading folder contents', error);
      documentGrid.innerHTML = '<p>Error loading folder contents.</p>';
    }
  }
}