/* eslint-disable @typescript-eslint/no-explicit-any */
import { MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Defines methods for interacting with Microsoft Graph.
 */
export interface IGraphService {
  getDriveItems(driveId: string, folderPath?: string): Promise<any>;
  searchDrive(query: string): Promise<any>;
  getListItemFields(driveId: string, itemId: string): Promise<any>;
  updateListItemFields(driveId: string, itemId: string, fields: any): Promise<any>;
  getSiteId(): Promise<string>;
  getDriveId(siteId: string): Promise<string>;
  listSensitivityLabels(): Promise<any[]>;
  getSensitivityLabel(driveId: string, itemId: string): Promise<{ id?: string; name?: string }>;
  assignSensitivityLabel(driveId: string, itemId: string, labelId: string, justification?: string): Promise<void>;
  listRetentionLabels(): Promise<any[]>;
  getRetentionLabel(driveId: string, itemId: string): Promise<any | undefined>;
  assignRetentionLabel(driveId: string, itemId: string, labelId: string): Promise<void>;
  evaluateDlp(driveId: string, itemId: string): Promise<{ isCompliant: boolean; violations?: any[] }>;
  getCustomerKeyStatus(): Promise<{ status: string; nextRotation?: string }>;
  listConditionalAccessPolicies(): Promise<any[]>;
  listInformationBarriers(): Promise<any[]>;
  uploadFileToDrive(drivePath: string, file: File): Promise<void>;
  getFolders(driveId: string, folderPath?: string): Promise<any[]>;
  getSharedWithMe(): Promise<any[]>;
  getDocumentActivities(driveId: string, itemId: string, top?: number): Promise<any[]>;
  getDocumentVersions(driveId: string, itemId: string, top?: number): Promise<any[]>;
  copyItemToOneDrive(itemId: string): Promise<void>;
}

/**
 * GraphService wraps calls to MS Graph using SPFx's MSGraphClient.
 */
export class GraphService implements IGraphService {
  private graphClient: MSGraphClientV3;

  private _siteIdCache?: string;
  private _driveIdCache: Map<string, string> = new Map();

  private readonly dlpPolicyId: string = 'a91d1fa7-91fa-4c36-8626-613c01ea71ee';

  private async fetchAllPages(request: any): Promise<any[]> {
    let result = await request.get();
    let items = result.value || [];
    while (result['@odata.nextLink']) {
      result = await this.graphClient.api(result['@odata.nextLink']).get();
      items = items.concat(result.value || []);
    }
    return items;
  }

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  /**
   * Retrieves the children of a folder in a drive.
   * @param driveId The ID of the drive.
   * @param folderPath Optional path in the drive.
   */
  public async getDriveItems(driveId: string, folderPath: string = ''): Promise<any> {
    const encodedPath = folderPath
      ? folderPath.split('/').map(encodeURIComponent).join('/')
      : '';
    const pathSegment = encodedPath
      ? `/root:/${encodedPath}:/children`
      : '/root/children';
    try {
      const request = this.graphClient.api(`/drives/${driveId}${pathSegment}`);
      return this.fetchAllPages(request);
    } catch (err) {
      throw new Error(`getDriveItems failed for drive ${driveId}: ${err}`);
    }
  }

  /**
   * Performs a search across the user's drive.
   * @param query The search query.
   */
  public async searchDrive(query: string): Promise<any[]> {
    try {
      const encodedQuery = encodeURIComponent(query);
      const request = this.graphClient.api(`/me/drive/root/search(q='${encodedQuery}')`);
      return this.fetchAllPages(request);
    } catch (err) {
      throw new Error(`searchDrive failed for query "${query}": ${err}`);
    }
  }

  /**
   * Retrieves the list item fields for a drive item.
   * @param driveId The ID of the drive.
   * @param itemId The ID of the item.
   */
  public getListItemFields(driveId: string, itemId: string): Promise<any> {
    return this.graphClient
      .api(`/drives/${driveId}/items/${itemId}/listItem/fields`)
      .get();
  }

  /**
   * Updates the list item fields for a drive item.
   * @param driveId The ID of the drive.
   * @param itemId The ID of the item.
   * @param fields An object containing the field values to update.
   */
  public updateListItemFields(driveId: string, itemId: string, fields: any): Promise<any> {
    return this.graphClient
      .api(`/drives/${driveId}/items/${itemId}/listItem/fields`)
      .patch(fields);
  }

  /**
   * Retrieves the SharePoint site ID for the root site.
   */
  public async getSiteId(): Promise<string> {
    if (this._siteIdCache) {
      return this._siteIdCache;
    }
    const site = await this.graphClient
      .api(`/sites/afribit.sharepoint.com:/sites/AFRIBIT432`)
      .get();
    this._siteIdCache = site.id;
    return site.id;
  }

  /**
   * Retrieves the default document library drive ID for a given site.
   * @param siteId The ID of the SharePoint site.
   */
  public async getDriveId(siteId: string): Promise<string> {
    if (this._driveIdCache.has(siteId)) {
      return this._driveIdCache.get(siteId)!;
    }
    const drive = await this.graphClient
      .api(`/sites/${siteId}/drive`)
      .get();
    this._driveIdCache.set(siteId, drive.id);
    return drive.id;
  }

  // Lists all published sensitivity labels in the tenant.
  public async listSensitivityLabels(): Promise<any[]> {
    const res = await this.graphClient.api('/dataClassification/sensitivityLabels').get();
    return res.value;
  }

  // Gets the current sensitivity label applied to a drive item.
  public async getSensitivityLabel(driveId: string, itemId: string): Promise<{ id?: string; name?: string }> {
    const fields = await this.getListItemFields(driveId, itemId);
    const tag = fields.complianceTag || fields.ComplianceTag;
    return tag ? { id: tag, name: tag } : {};
  }

  // Assigns a sensitivity label using the Graph action
  public async assignSensitivityLabel(driveId: string, itemId: string, labelId: string, justification: string = 'Assigned via Document Hub'): Promise<void> {
    await this.graphClient
      .api(`/drives/${driveId}/items/${itemId}/assignSensitivityLabel`)
      .version('beta')
      .post({
        labelId,
        justificationText: justification,
        assignmentMethod: 'standard'
      });
  }

  /**
   * Evaluates DLP policy compliance for a given file.
   * @param driveId The ID of the drive.
   * @param itemId The ID of the drive item.
   */
  public async evaluateDlp(driveId: string, itemId: string): Promise<{ isCompliant: boolean; violations?: any[] }> {
    try {
      // Get metadata with download URL
      const itemMeta: any = await this.graphClient
        .api(`/drives/${driveId}/items/${itemId}`)
        .select('@microsoft.graph.downloadUrl')
        .get();
      const downloadUrl: string = itemMeta['@microsoft.graph.downloadUrl'];
      console.log('[DLP] downloadUrl:', downloadUrl);
      if (!downloadUrl) {
        throw new Error(`No download URL available for item ${itemId}`);
      }
      // Download via fetch
      const fileResponse = await fetch(downloadUrl);
      const arrayBuffer: ArrayBuffer = await fileResponse.arrayBuffer();
      // Convert bytes to a binary string
      const binary = Array.from(new Uint8Array(arrayBuffer))
        .map(byte => String.fromCharCode(byte))
        .join('');
      const base64 = btoa(binary);
      console.log('[DLP] file content base64 size:', base64.length);
      // Call the DLP evaluation API
      console.log('[DLP] Calling evaluate endpoint for policyId:', this.dlpPolicyId);
      const response: any = await this.graphClient
        .api(`/security/informationProtection/dataLossPreventionPolicies/${this.dlpPolicyId}/evaluate`)
        .version('beta')
        .post({
          contentInfo: {
            source: 'file',
            fileContent: base64
          }
        });
      console.log('[DLP] API response:', response);
      return {
        isCompliant: response.evaluationResults?.[0]?.action === 'compliant',
        violations: response.evaluationResults?.[0]?.matchLocations || []
      };
    } catch (err) {
      console.error('[DLP] evaluateDlp error for item', itemId, err);
      throw new Error(`evaluateDlp failed for item ${itemId}: ${err}`);
    }
  }

  // Lists all published retention labels in the tenant.
  public async listRetentionLabels(): Promise<any[]> {
    const res = await this.graphClient.api('/security/retentionLabels').get();
    return res.value;
  }

  // Gets the current retention label applied to a drive item.
  public async getRetentionLabel(driveId: string, itemId: string): Promise<{ id?: string; name?: string } | undefined> {
    try {
      // Use v1.0 select on retentionLabel property instead of unsupported action endpoint
      const res: any = await this.graphClient
        .api(`/drives/${driveId}/items/${itemId}?$select=retentionLabel`)
        .get();
      const label = res.retentionLabel;
      return label ? { id: label.id, name: label.label } : undefined;
    } catch {
      return undefined;
    }
  }

  // Assigns a retention label to a drive item.
  public async assignRetentionLabel(driveId: string, itemId: string, labelId: string): Promise<void> {
    await this.graphClient
      .api(`/drives/${driveId}/items/${itemId}/setRetentionLabel`)
      .post({ labelId });
  }

  // Retrieves Customer-Managed Key (CMK) status.
  public async getCustomerKeyStatus(): Promise<{ status: string; nextRotation?: string }> {
    // TODO: connect to Purview / Key Vault management APIs
    return { status: 'Not configured', nextRotation: undefined };
  }

  // Lists Conditional Access policies for the tenant.
  public async listConditionalAccessPolicies(): Promise<any[]> {
    try {
      const request = this.graphClient.api('/identity/conditionalAccess/policies');
      return this.fetchAllPages(request);
    } catch (err) {
      throw new Error(`listConditionalAccessPolicies failed: ${err}`);
    }
  }

  // Lists Information Barriers policies.
  public async listInformationBarriers(): Promise<any[]> {
    try {
      // Call the correct beta endpoint for information barriers policies
      const request = this.graphClient
        .api('/beta/policies/authorizationPolicy') // adjust if necessary for actual information barriers endpoint
        .version('beta');
      // If the real path is /informationProtection/policy/informationBarriers, use:
      // const request = this.graphClient.api('/beta/informationProtection/policy/informationBarriers');
      const result: any = await request.get();
      return result.value || [];
    } catch {
      return [];
    }
  }
  /**
   * Test Graph connectivity by fetching the signed-in user profile.
   */
  public async testGraphCall(): Promise<any> {
    try {
      return await this.graphClient
        .api('/me')
        .get();
    } catch (err) {
      throw new Error(`testGraphCall failed: ${err}`);
    }
  }

  /** Uploads a file to the signed‑in user's OneDrive. */
  public async uploadFileToDrive(drivePath: string, file: File): Promise<void> {
    await this.graphClient
      .api(`/me/drive${drivePath}:/content`)
      .put(file);
  }

  /** Lists folders (only) under a given path in a drive. */
  public async getFolders(driveId: string, folderPath: string = ''): Promise<any[]> {
    const encodedPath = folderPath
      ? folderPath.split('/').map(encodeURIComponent).join('/')
      : '';
    const pathSegment = encodedPath
      ? `/root:/${encodedPath}:/children?$filter=folder ne null`
      : '/root/children?$filter=folder ne null';

    const res = await this.graphClient
      .api(`/drives/${driveId}${pathSegment}`)
      .get();
    return res.value;
  }

  /** Returns the items shared with the signed‑in user. */
  public async getSharedWithMe(): Promise<any[]> {
    const res = await this.graphClient.api('/me/drive/sharedWithMe').get();
    return res.value;
  }

  /** Gets recent activities for a drive item. */
  public async getDocumentActivities(driveId: string, itemId: string, top: number = 5): Promise<any[]> {
    const res = await this.graphClient
      .api(`/drives/${driveId}/items/${itemId}/activities?$top=${top}`)
      .get();
    return res.value;
  }

  /** Gets version history for a drive item. */
  public async getDocumentVersions(driveId: string, itemId: string, top: number = 5): Promise<any[]> {
    const res = await this.graphClient
      .api(`/drives/${driveId}/items/${itemId}/versions?$top=${top}`)
      .get();
    return res.value;
  }

  /** Copies a drive item to the root of the signed‑in user's OneDrive. */
  public async copyItemToOneDrive(itemId: string): Promise<void> {
    await this.graphClient
      .api(`/me/drive/items/${itemId}/copy`)
      .post({ parentReference: { id: 'root' } });
  }
}