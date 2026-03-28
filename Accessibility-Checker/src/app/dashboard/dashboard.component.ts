// dashboard.component.ts

import { Component } from '@angular/core';
import { environment } from '../../environments/environment';
const API_URL = environment.apiUrl;
import {
  HttpClient,
  HttpClientModule,
  HttpEvent,
  HttpEventType,
  HttpResponse,
} from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { FileUploadComponent } from '../file-upload/file-upload.component';
import { CommonModule } from '@angular/common';
import { HelpModalComponent } from '../help-modal/help-modal.component';

type RemediationIssue =
  | { type: 'fixed'; message: string }
  | { type: 'flagged'; message: string };

interface DocxRemediationResponse {
  fileName: string;
  suggestedFileName?: string | null;
  report: {
    fileName?: string;
    suggestedFileName?: string | null;
    summary: { fixed: number; flagged: number };
    details: {
      // FIXES (low risk)
      removedProtection?: boolean;
      // New remediation flags (backend)
      textShadowsRemoved?: boolean | number; // true or count
      fontsNormalized?: boolean | { replaced?: number };
      fontSizesNormalized?: boolean | { adjustedRuns?: number };
      minFontSizeEnforced?: boolean | { adjustedRuns?: number };
      lineSpacingFixed?: boolean | { adjustedParagraphs?: number };
      documentProtected?: boolean;
      fileNameFixed?: boolean;
      tablesHeaderRowSet?: Array<{ tableIndex: number }>;
      languageDefaultFixed?: { setTo: string };

      // DETECT-ONLY (names changed)
      fileNameNeedsFixing: boolean;
      titleNeedsFixing?: boolean;
      lineSpacingNeedsFixing?: boolean;
      fontTypeNeedsFixing?: boolean;
      fontSizeNeedsFixing?: boolean;
      linkNamesNeedImprovement?: boolean;
      formsDetected?: boolean;
      flashingObjectsDetected?: boolean;
      inlineContentFixed?: boolean;
      inlineContentCount?: number;
      inlineContentExplanation?: string;
      inlineContentCategory?: string;
      // Location details are provided in separate arrays
      lineSpacingLocations?: Array<{
        type: string;
        currentSpacing: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
      }>;
      fontTypeLocations?: Array<{
        type: string;
        font: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
      }>;
      fontSizeLocations?: Array<{
        type: string;
        size: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
      }>;
      emptyHeadings?: Array<{ paragraphIndex: number }>;
      headingOrderIssues?: Array<{
        paragraphIndex: number;
        previousLevel: number;
        currentLevel: number;
      }>;
      mergedSplitEmptyCells?: Array<{
        tableIndex: number;
        row: number;
        col: number;
        gridSpan?: number;
        vMerge?: any;
        isEmpty: boolean;
      }>;
      badLinks?: Array<{
        paragraphIndex: number;
        display: string;
        target?: string;
      }>;
      headerFooterAudit?: Array<{ part: string; preview: string }>;
      imagesMissingOrBadAlt?: number | Array<{
        imageIndex: number;
        page?: number;
        section?: string;
        description?: string;
        altText?: string;
      }>;
            // PowerPoint: auto-fixed alt text details (added by backend)
      autoFixedAltText?: Array<{
        slideNumber: number;
        shapeId?: string;
        fix: string;
        shapeName?: string;
        altText: string;
        note?: string;
      }>;
      // Image locations provided in separate array  
      imageLocations?: Array<{
        location: string;
        approximatePage: number;
        context: string;
        imagePath?: string;
        preview?: string;
        altText?: string;
      }>;
      // Link locations provided in separate array
      linkLocations?: Array<{
        type: string;
        linkText: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
        recommendation: string;
      }>;
      // Form locations provided in separate array
      formLocations?: Array<{
        type: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
        recommendation: string;
      }>;
      // Flashing object locations provided in separate array
      flashingObjectLocations?: Array<{
        type: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
        recommendation: string;
      }>;
      anchoredDrawingsDetected?: number;
      embeddedMedia?: Array<{ id: string; target: string; type: string }>;
      gifsDetected?: string[];
      // GIF/Flashing object locations provided in separate array
      gifLocations?: Array<{
        type: string;
        location: string;
        approximatePage: number;
        context: string;
        preview: string;
        recommendation: string;
        filePath?: string;
      }>;
      // Inline content positioning details
      inlineContentDetails?: Array<{
        type: string;
        description: string;
        originalContent?: string;
        fixedContent?: string;
      }>;
      colorContrastIssues?: Array<{
        paragraphIndex: number;
        color: string;
        sizePt?: number;
        bold?: boolean;
        ratio: number;
        sample: string;
      }>;
      languageDefaultIssue?: {
        current: string | null;
        recommendation: string;
      } | null;
      filenameFlag?: string | null;
    };
  };
}

interface ProcessedReport {
  response: DocxRemediationResponse;
  original?: File;
}

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [FileUploadComponent, HttpClientModule, CommonModule, HelpModalComponent],
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.css'],
})
export class DashboardComponent {
  // temporary debug flag to help diagnose "empty" dashboard issues
  // whether to show the unblock help modal after download
  showHelpModal = false;
  // show a small post-download banner with alternatives and a button to open modal
  showPostDownloadBanner = false;
  // debug: show raw remediation.report JSON
  showRawReport = false;
  isUploading = false;
  isServerProcessing = false;
  progress = 0;
  progressStageLabel = 'Ready';
  processingElapsedSeconds = 0;
  private processingExpectedSeconds = 30;
  private processingTimer?: ReturnType<typeof setInterval>;
  selectedFile?: File;
  fileName: string = '';
  downloadFileName: string = '';

  // backend response
  remediation?: DocxRemediationResponse;

  // store processed reports for multi-file batches so user can inspect each
  processedReports: ProcessedReport[] = [];

  // flattened list for the UI
  issues: RemediationIssue[] = [];

  // For debug/UX: list of auto-fixed items computed on the client
  getAutoFixedItems(): string[] {
    if (!this.remediation?.report?.details) return [];
    const d = this.remediation.report.details;
    const items: string[] = [];

    if (d.removedProtection) items.push('Document protection removed');
    if (d.fileNameFixed) items.push('File name fixed');
    if (d.tablesHeaderRowSet?.length) items.push(`${d.tablesHeaderRowSet.length} table header(s) set`);
    if (d.languageDefaultFixed) items.push(`Language set to ${d.languageDefaultFixed.setTo}`);
    // new backend flags
    const tsCount = this.getTextShadowsCount(d);
    if (tsCount > 0) {
      if (tsCount === 1) items.push('Text shadows removed');
      else items.push(`${tsCount} text shadow(s) removed`);
    }
    
    // Only show font size normalization as fixed (not font type or line spacing)
    if (d.fontSizesNormalized) {
      items.push('Font sizes normalized for consistency');
    }
    
    const minFontMsg = this.getMinFontSizeMessage(d);
    if (minFontMsg) items.push(minFontMsg);
    
    // Note: Line spacing and font types are now flagged for user action, not auto-fixed

    return items;
  }

  // Normalize the backend's textShadowsRemoved flag into a non-negative integer count
  private getTextShadowsCount(details: DocxRemediationResponse['report']['details'] | undefined): number {
    if (!details) return 0;
    const v = details.textShadowsRemoved;
    if (v === true) return 1; // boolean true means at least one was removed
    if (typeof v === 'number' && isFinite(v) && v > 0) return Math.floor(v);
    return 0;
  }

  // Return a user-friendly message for font normalization when the backend reports it
  getFontNormalizationMessage(details: DocxRemediationResponse['report']['details'] | undefined): string | null {
    if (!details) return null;
    // Prefer fontSizesNormalized (more specific) — caller/template can use this to avoid duplicates
    if (details.fontSizesNormalized) {
      if (typeof details.fontSizesNormalized === 'object' && (details.fontSizesNormalized as any).adjustedRuns)
        return `${(details.fontSizesNormalized as any).adjustedRuns} font size run(s) normalized`;
      return 'Font sizes normalized for consistency';
    }
    // Font types are now flagged for user action, not auto-fixed
    return null;
  }

  // Build a human-friendly message for min font size enforcement (or adjustments)
  private getMinFontSizeMessage(details: DocxRemediationResponse['report']['details'] | undefined): string | null {
    if (!details || details.minFontSizeEnforced === undefined || details.minFontSizeEnforced === null) return null;
    const v = details.minFontSizeEnforced;
    if (typeof v === 'object') {
      const adj = (v as any).adjustedRuns;
      const enforcedPt = (v as any).enforcedSizePt || (v as any).targetPt || (v as any).minSizePt;
      const sizeText = enforcedPt ? `${enforcedPt}pt` : '11pt';
      if (typeof adj === 'number' && adj > 0) return `${adj} run(s) adjusted to min font size (${sizeText})`;
      return `Minimum font size enforced (${sizeText})`;
    }
    return 'Minimum font size enforced (11pt)';
  }

  // --- UPDATED BATCH UPLOAD FUNCTION ---
  // Handle multiple files submitted from the file picker
  handleFiles(files: File[]) {
    if (!files || !files.length) return;

    // 1. Enforce 10-file limit
    if (files.length > 10) {
      alert('You can only upload a maximum of 10 files at once.');
      return;
    }

    this.beginProgressTracking(files.length);
    this.remediation = undefined; // Clear current main view
    this.issues = []; // Clear current issues

    // 2. Create ONE form data object
    const formData = new FormData();
    const uploadUrl = `${environment.apiUrl}${environment.uploadEndpoint}`;

    // 3. Append ALL files to the same 'files' key
    // This creates the list that Python expects: files=[f1, f2, f3]
    for (const f of files) {
      formData.append('files', f);
    }

    // 4. Send ONE Batch Request
    this.http
      .post(uploadUrl, formData, { observe: 'events', reportProgress: true })
      .subscribe({
        next: (event: HttpEvent<any>) => {
          if (event.type === HttpEventType.UploadProgress) {
            this.updateUploadProgress(event.loaded, event.total);
          } else if (event.type === HttpEventType.Response) {
            this.stopProcessingTimer();
            this.progress = 100;
            this.progressStageLabel = 'Completed';
            const body = (event as HttpResponse<any>).body;

            // Backend may return either { files: [...] } or { results: [...] }.
            const batchItems = Array.isArray(body?.files)
              ? body.files
              : Array.isArray(body?.results)
              ? body.results
              : [];

            if (batchItems.length) {
              const reportItems = batchItems.filter((res: any) => !!res?.report);
              const errorItems = batchItems.filter((res: any) => !!res?.error);

              const newReports: ProcessedReport[] = reportItems.map((res: any) => {
                const orig = files.find((f) => f.name === res.fileName) || files[0];
                return { response: res, original: orig };
              });

              if (newReports.length > 0) {
                const firstNewIndex = this.processedReports.length;
                this.processedReports.push(...newReports);
                this.selectReport(firstNewIndex);
              }

              if (errorItems.length > 0) {
                const errorMessages = errorItems.map(
                  (e: any) => `${e.fileName || 'Unknown file'}: ${e.error}`
                );
                this.issues = errorMessages.map((message: string) => ({
                  type: 'flagged',
                  message,
                }));
              }
            } else if (body && body.report) {
              // Fallback for single-object response.
              this.remediation = body;
              this.fileName = body.suggestedFileName || 'remediated.pptx';
              this.issues = this.flattenIssues(body);
              this.processedReports.push({ response: body, original: files[0] });
            } else {
              this.issues = [
                {
                  type: 'flagged',
                  message: 'Upload finished but server response had no readable report payload.',
                },
              ];
            }

            this.isUploading = false;
          }
        },
        error: (err) => {
          console.error('Batch Upload error:', err);
          this.stopProcessingTimer();
          this.progressStageLabel = 'Upload failed';
          this.isUploading = false;
          // Show error message (e.g. "Too many files" from backend)
          this.issues = [
            {
              type: 'flagged',
              message: `Upload failed: ${err?.error?.detail || err?.error?.message || err.statusText || 'Unknown error'}`,
            },
          ];
        },
      });
  }

  get rawReportJson(): string {
    try {
      return this.remediation?.report ? JSON.stringify(this.remediation.report, null, 2) : '';
    } catch (e) {
      return '';
    }
  }

  constructor(private http: HttpClient) {
    // debug logging removed
  }

  // Select a processed report to view its details in the main panel
  selectReport(index: number) {
    const rep = this.processedReports[index];
    if (!rep) return;
    // rep is a ProcessedReport { response, original }
    this.remediation = rep.response;
    this.issues = this.flattenIssues(rep.response);
    this.fileName = rep.response.suggestedFileName || rep.response.report?.fileName || '';
    // expose the original file so Download uses it
    this.selectedFile = rep.original;
    // reset download filename when switching
    this.downloadFileName = '';
  }

  // Download a specific processed report by index (uses stored original file when available)
  downloadReport(index: number) {
    const rep = this.processedReports[index];
    if (!rep) return;
    const targetFileName = rep.response?.suggestedFileName;

    if (!targetFileName) {
      this.issues = [{ type: 'flagged', message: 'No remediated file is available for this report.' }];
      return;
    }

    this.downloadFixed(targetFileName);
  }

  // --- UPDATED HANDLE FILE TO ALLOW POWERPOINT ---
  handleFile(payload: { file: File; title: string }) {
    const file = payload.file;
    const title = (payload.title || '').trim();
    this.selectedFile = file;
    this.beginProgressTracking(1);
    this.remediation = undefined;
    this.issues = [];

    // Check extensions
    const ext = file.name.split('.').pop()?.toLowerCase();
    const allowedExtensions = ['docx', 'pptx', 'ppt', 'pps', 'potx'];

    if (!ext || !allowedExtensions.includes(ext)) {
      this.isUploading = false;
      this.issues = [
        {
          type: 'flagged',
          message: `Unsupported file type ".${ext}". Please upload a PowerPoint (.pptx) or Word (.docx) file.`,
        },
      ];
      return;
    }

    // Directly upload without pinging the API
    this.uploadDocxFile(file, title);
  }

  private uploadDocxFile(file: File, title: string) {
    const uploadUrl = `${environment.apiUrl}${environment.uploadEndpoint}`;
    const formData = new FormData();
    // --- CHANGED: 'file' -> 'files' to match Python backend ---
    formData.append('files', file);
    formData.append('title', title); // backend may use this to set core title

    this.http
      .post(uploadUrl, formData, { observe: 'events', reportProgress: true })
      .subscribe({
        next: (event: HttpEvent<any>) => {
          if (event.type === HttpEventType.UploadProgress) {
            this.updateUploadProgress(event.loaded, event.total);
          // } else if (event.type === HttpEventType.Response) {
          //   const res = (event as HttpResponse<any>)
          //     .body as DocxRemediationResponse;
          //   this.remediation = res;
          //   this.fileName = res.suggestedFileName ? res.suggestedFileName : "remediated.pptx";
          //   this.issues = this.flattenIssues(res);

          //   // Save processed report so the user can inspect it individually later (include original file)
          //   try {
          //     if (res && res.report) this.processedReports.push({ response: res, original: file });
          //   } catch (e) {}
          //   this.isUploading = false;
          // }
          } else if (event.type === HttpEventType.Response) {
            this.stopProcessingTimer();
            this.progress = 100;
            this.progressStageLabel = 'Completed';
            const body = (event as HttpResponse<any>).body;

            const res =
              body?.files && Array.isArray(body.files) && body.files.length
                ? body.files[0]
                : body;

            if (!res || !res.report) {
              this.isUploading = false;
              this.issues = [
                { type: 'flagged', message: 'Unexpected upload response from server.' },
              ];
              return;
            }

            this.remediation = res;
            this.fileName = res.suggestedFileName || 'remediated.pptx';
            this.issues = this.flattenIssues(res);

            try {
              this.processedReports.push({ response: res, original: file });
            } catch {}

            this.isUploading = false;
          }
        },
        error: (err) => {
          console.error('Upload error:', err);
          this.stopProcessingTimer();
          this.progressStageLabel = 'Upload failed';
          this.isUploading = false;
          this.issues = [
            {
              type: 'flagged',
              message: `Upload failed: ${err?.error?.message || err.statusText || err.message || 'Unknown error'}`,
            },
          ];
        },
      });
  }

  handleClear() {
    this.stopProcessingTimer();
    this.selectedFile = undefined;
    this.remediation = undefined;
    this.issues = [];
    this.isUploading = false;
    this.progress = 0;
    this.progressStageLabel = 'Ready';
    this.fileName = '';
    this.downloadFileName = '';
    this.processedReports = []; // Clear all processed reports
  }

  private startProcessingTimer() {
    if (this.isServerProcessing) return;
    this.isServerProcessing = true;
    this.progressStageLabel = 'Validating uploaded files...';
    this.progress = Math.max(this.progress, 40);
    this.processingElapsedSeconds = 0;
    this.processingTimer = setInterval(() => {
      this.processingElapsedSeconds += 1;

      // Move smoothly from 40% to 95% based on estimated server work.
      const normalized = Math.min(
        1,
        this.processingElapsedSeconds / this.processingExpectedSeconds,
      );
      const stagedProgress = 40 + Math.floor(normalized * 55);
      this.progress = Math.max(this.progress, Math.min(95, stagedProgress));

      if (this.progress < 55) {
        this.progressStageLabel = 'Validating uploaded files...';
      } else if (this.progress < 80) {
        this.progressStageLabel = 'Running AI remediation...';
      } else if (this.progress < 92) {
        this.progressStageLabel = 'Building accessibility reports...';
      } else {
        this.progressStageLabel = 'Finalizing response...';
      }
    }, 1000);
  }

  private stopProcessingTimer() {
    this.isServerProcessing = false;
    if (this.processingTimer) {
      clearInterval(this.processingTimer);
      this.processingTimer = undefined;
    }
  }

  private beginProgressTracking(fileCount: number) {
    this.stopProcessingTimer();
    this.isUploading = true;
    this.isServerProcessing = false;
    this.progress = 1;
    this.progressStageLabel = 'Preparing files...';
    this.processingElapsedSeconds = 0;
    this.processingExpectedSeconds = Math.max(20, fileCount * 8);
  }

  private updateUploadProgress(loaded: number, total?: number) {
    this.progressStageLabel = 'Uploading files...';

    if (!total || total <= 0) {
      this.progress = Math.max(this.progress, 8);
      return;
    }

    // Reserve 0-40% for upload stage so processing can visibly occupy the rest.
    const uploadPercent = Math.min(100, (100 * loaded) / total);
    const mapped = Math.round(5 + (uploadPercent * 35) / 100);
    this.progress = Math.max(this.progress, Math.min(40, mapped));

    if (loaded >= total) {
      this.startProcessingTimer();
    }
  }

  formatElapsed(seconds: number): string {
    const m = Math.floor(seconds / 60)
      .toString()
      .padStart(2, '0');
    const s = (seconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
  }

  // Remove individual report
  removeReport(index: number) {
    this.processedReports.splice(index, 1);
    
    // If we removed the currently selected report, clear the view
    if (this.remediation && this.processedReports.length === 0) {
      this.remediation = undefined;
      this.issues = [];
    }
  }

  private flattenIssues(res: DocxRemediationResponse): RemediationIssue[] {
  
    if (!res?.report?.details) return [];
    const d = res.report.details;
    const out: RemediationIssue[] = [];

    // FIXED (conservative)
    if (d.removedProtection)
      out.push({
        type: 'fixed',
        message:
          'Document protection has been successfully removed, allowing full editing access.',
      });
    // New backend fixes
    const tsCountFlat = this.getTextShadowsCount(d);
    if (tsCountFlat > 0)
      out.push({
        type: 'fixed',
        message:
          tsCountFlat === 1
            ? 'Text shadows were removed to improve text legibility.'
            : `${tsCountFlat} text shadow style(s) were removed for improved readability.`,
      });

    // Font types are now flagged for user action, not auto-fixed
    
    if (d.fontSizesNormalized) {
      out.push({
        type: 'fixed',
        message:
          typeof d.fontSizesNormalized === 'object' && (d.fontSizesNormalized as any).adjustedRuns
            ? `${(d.fontSizesNormalized as any).adjustedRuns} font size run(s) were normalized for consistency.`
            : 'Font sizes were normalized for consistency.',
      });
    }

    const minFontMsgFlat = this.getMinFontSizeMessage(d);
    if (minFontMsgFlat) out.push({ type: 'fixed', message: minFontMsgFlat.includes('adjusted') ? minFontMsgFlat.replace('adjusted to min font size', 'enforced to') : minFontMsgFlat.replace('Minimum font size enforced', 'Minimum font size enforced to') });
    
    // Inline content positioning fixes
    if (d.inlineContentFixed && d.inlineContentCount) {
      let message = d.inlineContentExplanation || 'Inline content positioning fixed for better accessibility.';
      
      if (d.inlineContentCount > 1) {
        message = `${d.inlineContentCount} inline content positioning issues fixed. ${d.inlineContentExplanation || 'In-line content is necessary for users who rely on assistive technology and preferred by all users.'}`;
      }
      
      // Add category information if available
      if (d.inlineContentCategory) {
        message += ` (${d.inlineContentCategory})`;
      }
      
      // If detailed information is available, include it
      if (Array.isArray(d.inlineContentDetails) && d.inlineContentDetails.length > 0) {
        message += `\n\n📊 Details - Click to expand`;
        
        const detailsList = d.inlineContentDetails.map((item, index) => {
          let detail = `${index + 1}. ${item.description || 'Positioning fix applied'}`;
          if (item.type) detail += ` (${item.type})`;
          if (item.originalContent && typeof item.originalContent === 'string' && !item.originalContent.includes('<')) {
            detail += `\n    Original: "${item.originalContent.substring(0, 60)}..."`;
          }
          return detail;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${d.inlineContentDetails.length} Fix Detail${d.inlineContentDetails.length > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${detailsList}</div>\n</details>`;
      }
      
      out.push({
        type: 'fixed',
        message: message,
      });
    }
    
    // Line spacing and font type are now flagged for user action
    if (d.lineSpacingNeedsFixing) {
      let message = 'Line spacing needs to be at least 1.5 for improved readability (flagged for your attention).';
      
      // If detailed location information is available, include it
      if (d.lineSpacingLocations && d.lineSpacingLocations.length > 0) {
        const count = d.lineSpacingLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = d.lineSpacingLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.currentSpacing) location += ` • Current: ${item.currentSpacing}`;
          if (item.preview && !item.preview.includes('<w:')) {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} Line Spacing Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }
    
    if (d.fontTypeNeedsFixing) {
      out.push({
        type: 'flagged',
        message: 'Font types should be Arial/sans-serif for better accessibility (flagged for your attention).',
      });
    }
    
    if (d.fontSizeNeedsFixing) {
      let message = 'Font sizes should be consistent and appropriately sized (flagged for your attention).';
      
      // If detailed location information is available, include it
      if (d.fontSizeLocations && d.fontSizeLocations.length > 0) {
        const count = d.fontSizeLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = d.fontSizeLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.size) location += ` • Current: ${item.size}`;
          if (item.preview && !item.preview.includes('<w:')) {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} Font Size Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }
    
    if (d.linkNamesNeedImprovement) {
      let message = 'Link names are not descriptive (flagged for your attention).';
      
      // If detailed location information is available, include it
      if (d.linkLocations && d.linkLocations.length > 0) {
        // Remove duplicates based on location and linkText
        const uniqueLocations = d.linkLocations.filter((item, index, self) =>
          index === self.findIndex(t => t.location === item.location && t.linkText === item.linkText)
        );
        
        const count = uniqueLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = uniqueLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.type) location += ` • Type: ${item.type}`;
          if (item.linkText) location += ` • Current text: "${item.linkText}"`;
          if (item.recommendation) location += ` • Recommendation: ${item.recommendation}`;
          if (item.preview && !item.preview.includes('<w:')) {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} Link Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }
    
    if (d.formsDetected) {
      let message = 'Forms detected (flagged for your attention).';
      
      // If detailed location information is available, include it
      if (d.formLocations && d.formLocations.length > 0) {
        const count = d.formLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = d.formLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.type) location += ` • Type: ${item.type}`;
          if (item.recommendation) location += ` • Recommendation: ${item.recommendation}`;
          if (item.preview && !item.preview.includes('<w:')) {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} Form Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }
    
    if (d.flashingObjectsDetected) {
      let message = 'Flashing objects detected (flagged for your attention).';
      
      // If detailed location information is available, include it
      if (d.flashingObjectLocations && d.flashingObjectLocations.length > 0) {
        const count = d.flashingObjectLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = d.flashingObjectLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.type) location += ` • Type: ${item.type}`;
          if (item.recommendation) location += ` • Recommendation: ${item.recommendation}`;
          if (item.preview && !item.preview.includes('<w:')) {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} Flashing Object Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }
    
    if (d.documentProtected === true)
      out.push({
        type: 'fixed',
        message:
          'Document protection was removed to allow full editing access.',
      });

    if (d.fileNameFixed)
      out.push({
        type: 'fixed',
        message:
          'The file name has been updated to a more descriptive and accessible name.',
      });
    if (d.titleNeedsFixing)
      out.push({
        type: 'flagged',
        message:
          'The document title needs to be updated. It should be set to something descriptive.',
      });
    if (d.tablesHeaderRowSet?.length)
      out.push({
        type: 'fixed',
        message: `Repeating header row has been set for ${d.tablesHeaderRowSet.length} table(s) for better accessibility in long documents.`,
      });
    if (d.languageDefaultFixed)
      out.push({
        type: 'fixed',
        message: `The document language has been set to ${d.languageDefaultFixed.setTo} for consistent language and readability.`,
      });

    // FLAGGED (detect-only)
    if (d.fileNameNeedsFixing) {
      out.push({
        type: 'flagged',
        message: `The file name is too generic or unclear. Please make sure it reflects the content of the document and avoids terms like 'document' or 'untitled'. Also, replace any underscores (_) with hyphens (-) for better readability.`,
      });
    }

    if (d.emptyHeadings?.length)
      out.push({
        type: 'flagged',
        message: `${d.emptyHeadings.length} empty heading(s) detected. Headings should contain meaningful text for structure and accessibility.`,
      });

    if (d.headingOrderIssues?.length)
      out.push({
        type: 'flagged',
        message: `${d.headingOrderIssues.length} heading order issue(s) detected. Ensure headings are in proper order (e.g., Heading 1 followed by Heading 2).`,
      });

    if (d.badLinks?.length)
      out.push({
        type: 'flagged',
        message: `${d.badLinks.length} hyperlink(s) need descriptive text. Use clear, meaningful link text that indicates the target of the link.`,
      });

    if (d.mergedSplitEmptyCells?.length)
      out.push({
        type: 'flagged',
        message: `${d.mergedSplitEmptyCells.length} table cell issue(s) detected (merged, split, or empty). Ensure all table cells are properly structured for readability.`,
      });

    if (d.headerFooterAudit?.length)
      out.push({
        type: 'flagged',
        message: `The header/footer contains content. Consider duplicating key information within the document body for better accessibility.`,
      });

    // Handle PowerPoint format (array of objects with slideNumber)
    if (d.imagesMissingOrBadAlt && Array.isArray(d.imagesMissingOrBadAlt) && d.imagesMissingOrBadAlt.length > 0) {
      const count = d.imagesMissingOrBadAlt.length;
      
      const locationList = (d.imagesMissingOrBadAlt as any[]).map((item, index) => {
        const location = item.location || `Slide ${item.slideNumber}` || `Image ${index + 1}`;
        return `<li>${location}</li>`;
      }).join('');
      
      const message = `<strong>${count} image(s) are missing alt text</strong><br/>` +
        `<ul style="margin: 8px 0 0 20px; padding: 0; list-style-type: disc;">${locationList}</ul>`;
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }
    // Handle Word format (number or separate imageLocations array)
    else if (d.imagesMissingOrBadAlt && (typeof d.imagesMissingOrBadAlt === 'number' && d.imagesMissingOrBadAlt > 0)) {
      let message = `${d.imagesMissingOrBadAlt} image(s) are missing or have poor alt text. Alt text should describe the content and purpose of the image for accessibility.`;
      
      // If detailed location information is available, include it
      if (d.imageLocations && d.imageLocations.length > 0) {
        const count = d.imageLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = d.imageLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.imagePath) location += ` • File: ${item.imagePath}`;
          if (item.altText) location += ` • Current alt: "${item.altText}"`;
          if (item.preview && item.preview !== 'No surrounding text') {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} Image Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }

    if (
      typeof d.anchoredDrawingsDetected === 'number' &&
      d.anchoredDrawingsDetected > 0
    )
      out.push({
        type: 'flagged',
        message: `${d.anchoredDrawingsDetected} anchored drawing(s) detected. For improved reading order, use inline images instead.`,
      });

    if (d.embeddedMedia?.length)
      out.push({
        type: 'flagged',
        message: `${d.embeddedMedia.length} embedded media item(s) detected. Ensure captions or transcripts are provided for accessibility.`,
      });

    if (d.gifsDetected?.length) {
      let message = `${d.gifsDetected.length} GIF(s) detected. Verify that none of them contain flashing content, which can cause accessibility issues.`;
      
      // If detailed location information is available, include it
      if (d.gifLocations && d.gifLocations.length > 0) {
        const count = d.gifLocations.length;
        message += `\n\n📍 ${count} location${count > 1 ? 's' : ''} found - Click to expand details`;
        
        const locationDetails = d.gifLocations.map((item, index) => {
          let location = `${index + 1}. ${item.location}`;
          if (item.approximatePage) location += ` (Page ${item.approximatePage})`;
          if (item.context && item.context !== 'Document body') location += ` • ${item.context}`;
          if (item.type) location += ` • Type: ${item.type}`;
          if (item.filePath) location += ` • File: ${item.filePath}`;
          if (item.recommendation) location += ` • Recommendation: ${item.recommendation}`;
          if (item.preview && !item.preview.includes('<w:')) {
            location += `\n    Preview: "${item.preview.substring(0, 80)}..."`;
          }
          return location;
        }).join('\n\n');
        
        message += `\n\n<details class="mt-2">\n<summary class="cursor-pointer text-sm font-medium text-blue-600 dark:text-blue-400 hover:underline">View ${count} GIF Issue${count > 1 ? 's' : ''}</summary>\n<div class="mt-2 pl-4 text-sm text-gray-600 dark:text-gray-300 whitespace-pre-line">${locationDetails}</div>\n</details>`;
      }
      
      out.push({
        type: 'flagged',
        message: message,
      });
    }

    // --- PowerPoint: missing slide titles ---
  if ((d as any).slidesMissingTitles?.length) {
    for (const s of (d as any).slidesMissingTitles) {
      const slideNum = s?.slideNumber ?? '?';
      const msg = s?.message || `Slide ${slideNum} is missing a title`;
      out.push({
        type: 'flagged',
        message: msg,
      });
    }
  }

    if (d.colorContrastIssues?.length)
      out.push({
        type: 'flagged',
        message: `${d.colorContrastIssues.length} low contrast text run(s) detected against a white background. Consider increasing contrast for readability.`,
      });

    if (d.languageDefaultIssue && !d.languageDefaultFixed)
      out.push({
        type: 'flagged',
        message: `Document default language is set to ${d.languageDefaultIssue.current || 'unset'}. It is recommended to set it to ${d.languageDefaultIssue.recommendation} for consistency.`,
      });

    if (d.filenameFlag) out.push({ type: 'flagged', message: d.filenameFlag });

    return out;
  }

  // downloadFixed() {
  //   const downloadUrl = `${environment.apiUrl}${environment.downloadEndpoint}`;

  //   if (!this.selectedFile) {
  //     console.error('No file selected for download');
  //     return; // Early return if no file is selected
  //   }
  //   // Prepare form data to send file to the server
  //   const formData = new FormData();
  //   formData.append('file', this.selectedFile); // Add the file object

  //   // Send POST request with file object
  //   this.http
  //     .post(downloadUrl, formData, {
  //       responseType: 'blob',
  //       observe: 'response',
  //     })
  //     .subscribe({
  //       next: (response: HttpResponse<Blob>) => {
  //         // Check if the response body is not null
  //         const blob = response.body;
  //         if (!blob) {
  //           console.error('Error: Empty response body');
  //           return;
  //         }

  //         const contentType = (response.headers.get('content-type') || '').toLowerCase();

  //         // If server returned JSON (error payload), parse and show a user message
  //         if (contentType.includes('application/json')) {
  //           // blob.text() returns a promise with the JSON string
  //           blob.text().then((txt) => {
  //             try {
  //               const payload = JSON.parse(txt);
  //               this.issues = [
  //                 { type: 'flagged', message: payload?.error || 'Server error during remediation' },
  //               ];
  //             } catch (e) {
  //               this.issues = [
  //                 { type: 'flagged', message: 'Unexpected server response during remediation.' },
  //               ];
  //             }
  //           });
  //           return;
  //         }

  //         // Extract filename from Content-Disposition header (supports filename and filename*=)
  //         const contentDisposition = response.headers.get('content-disposition') || response.headers.get('Content-Disposition') || '';
  //         let filename = 'remediated-document.docx'; // default

  //         if (contentDisposition) {
  //           // Try filename*=UTF-8''name.docx first
  //           const fstar = contentDisposition.match(/filename\*=[^']*''([^;\n\r]+)/i);
  //           if (fstar && fstar[1]) {
  //             try {
  //               filename = decodeURIComponent(fstar[1]);
  //             } catch (e) {
  //               filename = fstar[1];
  //             }
  //           } else {
  //             const matches = /filename=\s*"?([^";]+)"?/i.exec(contentDisposition);
  //             if (matches && matches[1]) filename = matches[1];
  //           }
  //         }

  //         // Store the filename for display purposes
  //         this.downloadFileName = filename;

  //         // Note: do not mutate the server-provided summary here. The UI shows a
  //         // client-estimated auto-fix count with `countAutoFixableIssues()` and
  //         // we perform an authoritative re-check immediately after download that
  //         // replaces the report with the server's post-remediation response.

  //         // Create download link with the correct filename
  //         const url = URL.createObjectURL(blob);
  //         const a = document.createElement('a');
  //         a.href = url;
  //         a.download = filename; // Use the filename from the header
  //         a.click();
  //         URL.revokeObjectURL(url);

  //         // If the original report said the document was protected, show a post-download banner
  //         // with alternatives and a button to open the unblock help modal.
  //         if (this.remediation?.report?.details?.documentProtected) {
  //           this.showPostDownloadBanner = true;
  //         }
          
  //         const isPptx = filename.toLowerCase().endsWith('.pptx');
  //         if (isPptx) return;
  //         // Authoritative re-check: send the downloaded blob back to the upload analysis endpoint
  //         // so the UI shows the exact server-side post-remediation report (avoids client heuristics).
  //         try {
  //           const analysisUrl = `${environment.apiUrl}${environment.uploadEndpoint}`;
  //           const reForm = new FormData();
  //           // Append blob as a file; use the filename determined above so server sees correct name
  //           reForm.append('file', blob, filename);

  //           this.http.post(analysisUrl, reForm).subscribe({
  //             next: (resp: any) => {
  //               // Replace remediation with authoritative server response and re-render issues
  //               try {
  //                 const res = resp as DocxRemediationResponse;
  //                 if (res && res.report) {
  //                   this.remediation = res;
  //                   this.fileName = res.suggestedFileName ? res.suggestedFileName : filename;
  //                   this.issues = this.flattenIssues(res);
  //                   // Hide post-download banner if server confirms protection removed
  //                   if (!this.remediation.report.details?.documentProtected) {
  //                     this.showPostDownloadBanner = false;
  //                   }
  //                 }
  //               } catch (e) {
  //                 console.warn('Failed to parse authoritative report', e);
  //               }
  //             },
  //             error: (err) => {
  //               console.warn('Authoritative re-check failed', err);
  //               // keep existing remediation but surface a message
  //               this.issues = [
  //                 { type: 'flagged', message: `Could not refresh authoritative report: ${err?.message || err?.statusText || 'error'}` },
  //               ];
  //             },
  //           });
  //         } catch (e) {
  //           console.warn('Authoritative re-check error', e);
  //         }
  //       },
  //       error: (err) => {
  //         console.error('Download failed', err);
  //         this.issues = [
  //           { type: 'flagged', message: `Download failed: ${err?.error?.message || err.statusText || err.message || 'Unknown error'}` },
  //         ];
  //       },
  //     });
  // }

  downloadFixed(targetFileName?: string) {
    const downloadUrl = `${environment.apiUrl}${environment.downloadEndpoint}`;

    const fileName =
      targetFileName ||
      this.remediation?.suggestedFileName ||
      this.fileName ||
      this.downloadFileName;

    if (!fileName) {
      this.issues = [{ type: 'flagged', message: 'No remediated file available for download.' }];
      return;
    }

    this.http
      .post(
        downloadUrl,
        { fileName },
        {
          responseType: 'blob',
          observe: 'response',
        }
      )
      .subscribe({
        next: (response: HttpResponse<Blob>) => {
          const blob = response.body;
          if (!blob) {
            this.issues = [{ type: 'flagged', message: 'Empty download response.' }];
            return;
          }

          const contentDisposition =
            response.headers.get('content-disposition') ||
            response.headers.get('Content-Disposition') ||
            '';

          let filename = this.cleanName(fileName);

          if (contentDisposition) {
            const fstar = contentDisposition.match(/filename\*=[^']*''([^;\n\r]+)/i);
            if (fstar && fstar[1]) {
              try {
                filename = this.cleanName(decodeURIComponent(fstar[1]));
              } catch {
                filename = this.cleanName(fstar[1]);
              }
            } else {
              const matches = /filename=\s*"?([^";]+)"?/i.exec(contentDisposition);
              if (matches && matches[1]) filename = this.cleanName(matches[1]);
            }
          }

          this.downloadFileName = filename;

          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          // a.download = filename;
          a.download = this.cleanName(fileName);
          a.click();
          URL.revokeObjectURL(url);
        },
        error: (err) => {
          this.issues = [
            {
              type: 'flagged',
              message: `Download failed: ${err?.error?.detail || err?.message || 'Unknown error'}`,
            },
          ];
        },
      });
  }

  countAutoFixableIssues(): number {
    // Count issues that the backend automatically fixes during download/remediation
    // These are the same issues that show "fixed" status but weren't counted yet
    // ✅ PPTX path: just trust backend summary.fixed
    const details: any = this.remediation?.report?.details;
    const summary: any = this.remediation?.report?.summary;

    // PPTX reports have these keys
    if (details?.imagesMissingOrBadAlt || details?.slidesMissingTitles) {
      return summary?.fixed ?? 0;
    }
    if (!this.remediation?.report?.details) return 0;
    
    const d = this.remediation.report.details;
    let count = 0;
    
    // These are auto-fixed by the backend during remediation:
    if (d.documentProtected === true) count++; // Protection removal
    if (d.fileNameNeedsFixing && !d.fileNameFixed) count++; // Filename fix
    if (d.tablesHeaderRowSet?.length) count += d.tablesHeaderRowSet.length; // Table headers
    if (d.languageDefaultIssue && !d.languageDefaultFixed) count++; // Language fix
  // New backend auto-fixes
  const tsCount = this.getTextShadowsCount(d);
  if (tsCount > 0) count += tsCount;
    // Only count font size normalization (not font type or line spacing - those are now flagged)
    if (d.fontSizesNormalized) {
      if (typeof d.fontSizesNormalized === 'object' && (d.fontSizesNormalized as any).adjustedRuns)
        count += (d.fontSizesNormalized as any).adjustedRuns;
      else count++;
    }
    
    if (d.minFontSizeEnforced) {
      if (typeof d.minFontSizeEnforced === 'object' && d.minFontSizeEnforced.adjustedRuns)
        count += d.minFontSizeEnforced.adjustedRuns;
      else count++;
    }
    
    // Note: Line spacing and font types are now flagged for user action, not auto-fixed
    
    return count;
  }

  // Count actual flagged issues as displayed to the user (forms count as 1 regardless of locations)
  getFlaggedCount(): number {
    const details: any = this.remediation?.report?.details;
    const summary: any = this.remediation?.report?.summary;

    // PPTX: trust backend summary
    if (details?.imagesMissingOrBadAlt || details?.slidesMissingTitles) {
      return summary?.flagged ?? 0;
    }

    // DOCX: count rendered issues
    if (!this.issues) return 0;
    return this.issues.filter(issue => issue.type === 'flagged').length;
  }

  getFixedCount(): number {
    const details: any = this.remediation?.report?.details;
    const summary: any = this.remediation?.report?.summary;

    // PPTX: trust backend summary
    if (details?.imagesMissingOrBadAlt || details?.slidesMissingTitles) {
      return summary?.fixed ?? 0;
    }

    // DOCX: count rendered issues
    if (!this.issues) return 0;
    return this.issues.filter(issue => issue.type === 'fixed').length;
  }
  // getFlaggedCount(): number {
  //   if (!this.issues) return 0;
  //   return this.issues.filter(issue => issue.type === 'flagged').length;
  // }

  // // Count actual fixed issues as displayed to the user
  // getFixedCount(): number {
  //   if (!this.issues) return 0;
  //   return this.issues.filter(issue => issue.type === 'fixed').length;
  // }

  cleanName(name: string): string {
    return name.replace(/^[0-9a-f]{8}_/i, '');
  }

}