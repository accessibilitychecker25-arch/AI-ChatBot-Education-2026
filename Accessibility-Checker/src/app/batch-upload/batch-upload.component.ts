// import { Component, OnDestroy } from '@angular/core';
// import { CommonModule } from '@angular/common';
// import { BatchUploadService } from './batch-upload.service';
// import { HttpClientModule, HttpEventType } from '@angular/common/http';

// @Component({
//   selector: 'app-batch-upload',
//   standalone: true,
//   imports: [CommonModule, HttpClientModule],
//   templateUrl: './batch-upload.component.html',
//   styleUrls: ['./batch-upload.component.css'],
// })
// export class BatchUploadComponent implements OnDestroy {
//   sessionId?: string;
//   files: File[] = [];
//   uploadProgress = 0; // Overall upload progress (0-100%)
//   isProcessing = false; // True while files are being processed
//   results: any[] = [];
//   error?: string;
//   keepAliveInterval?: any;
  
//   // Track which file is currently being processed
//   currentFileIndex = -1;
//   processingStatus: { [key: number]: string } = {};

//   constructor(private svc: BatchUploadService) {}

//   createSession() {
//     this.error = undefined;
//     this.svc.createSession().subscribe({
//       next: (resp) => {
//         this.sessionId = resp?.sessionId;
//         // start keepalive
//         this.keepAliveInterval = setInterval(() => this.svc.keepAlive(this.sessionId!).subscribe(), 5 * 60 * 1000);
//       },
//       error: (e: any) => {
//         this.error = `Session creation failed: ${e?.message || e?.statusText || e}`;
//       },
//     });
//   }

//   onFilesSelected(ev: Event) {
//     const input = ev.target as HTMLInputElement;
//     if (!input.files || input.files.length === 0) return;

//     // Enforce 10-File Limit
//     if (input.files.length > 10) {
//       alert("You can only upload a maximum of 10 files at once.");
//       input.value = '';
//       this.files = [];
//       return;
//     }

//     const selectedFiles = Array.from(input.files);

//     // Enforce PowerPoint File Type
//     const validExtensions = /\.(pptx|ppt|pps|pot|potx|ppsx)$/i;

//     for (let file of selectedFiles) {
//       if (!validExtensions.test(file.name)) {
//         alert(`File "${file.name}" is not a valid PowerPoint file.`);
//         input.value = '';
//         this.files = [];
//         return;
//       }
//     }

//     this.files = selectedFiles;
//     this.error = undefined;
//   }

//   upload() {
//     this.error = undefined;
//     this.results = [];
//     this.uploadProgress = 0;
//     this.isProcessing = true;
//     this.processingStatus = {};
//     this.currentFileIndex = -1;
    
//     if (!this.files.length) {
//       this.error = 'No files selected';
//       this.isProcessing = false;
//       return;
//     }

//     if (this.files.length > 10) {
//       this.error = 'Too many files selected (Max 10)';
//       this.isProcessing = false;
//       return;
//     }

//     // Show initial status for all files
//     for (let i = 0; i < this.files.length; i++) {
//       this.processingStatus[i] = 'Queued';
//     }

//     this.svc.batchUpload(this.files, this.sessionId).subscribe({
//       next: (event: any) => {
//         if (event.type === HttpEventType.UploadProgress) {
//           // Show upload progress (before processing)
//           this.uploadProgress = Math.round((100 * event.loaded) / (event.total ?? 1));
//         } else if (event.type === HttpEventType.Response) {
//           const body = event.body || {};
          
//           // Update session
//           if (body.sessionId) this.sessionId = body.sessionId;
          
//           // Batch results now includes all file results
//           this.results = body.files || [];
          
//           // Update processing status for each result
//           for (let i = 0; i < this.results.length; i++) {
//             const result = this.results[i];
//             if (result.error) {
//               this.processingStatus[i] = `Error: ${result.error}`;
//             } else if (result.report) {
//               this.processingStatus[i] = `✅ Completed`;
//             }
//           }
          
//           this.isProcessing = false;
//           this.uploadProgress = 100;
//         }
//       },
//       error: (err: any) => {
//         this.isProcessing = false;
//         if (err?.error) this.error = JSON.stringify(err.error);
//         else this.error = err?.message || err?.statusText || 'Upload failed';
//       },
//     });
//   }

//   downloadFile(suggestedFileName: string) {
//     if (!suggestedFileName) return;
//     this.svc.downloadFile(suggestedFileName).subscribe({
//       next: (blob) => {
//         const url = URL.createObjectURL(blob);
//         const a = document.createElement('a');
//         a.href = url;
//         a.download = suggestedFileName;
//         a.click();
//         URL.revokeObjectURL(url);
//       },
//       error: (err: any) => {
//         this.error = `Download failed: ${err?.message || 'error'}`;
//       },
//     });
//   }

//   downloadAll() {
//     this.error = undefined;
//     if (!this.sessionId) {
//       this.error = 'No session available for download';
//       return;
//     }
//     this.svc.getBatchDownload(this.sessionId).subscribe({
//       next: (blob) => {
//         const filename = `remediated-files.zip`;
//         const url = URL.createObjectURL(blob);
//         const a = document.createElement('a');
//         a.href = url;
//         a.download = filename;
//         a.click();
//         URL.revokeObjectURL(url);
//       },
//       error: (err: any) => {
//         this.error = `Batch download failed: ${err?.message || err?.statusText || 'error'}`;
//       },
//     });
//   }

//   getTotalFixed(): number {
//     return this.results.reduce((sum, result) => {
//       return sum + (result.report?.summary?.fixed || 0);
//     }, 0);
//   }

//   hasDownloadableFiles(): boolean {
//     return this.results && this.results.length > 0 && 
//            this.results.some((r: any) => r.suggestedFileName && !r.error);
//   }

//   ngOnDestroy() {
//     if (this.keepAliveInterval) clearInterval(this.keepAliveInterval);
//   }
// }

import { Component, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { BatchUploadService } from './batch-upload.service';
import { HttpClientModule, HttpEventType } from '@angular/common/http';

@Component({
  selector: 'app-batch-upload',
  standalone: true,
  imports: [CommonModule, HttpClientModule],
  templateUrl: './batch-upload.component.html',
  styleUrls: ['./batch-upload.component.css'],
})
export class BatchUploadComponent implements OnDestroy {
  files: File[] = [];
  uploadProgress = 0;
  isProcessing = false;
  results: any[] = [];
  error?: string;

  currentFileIndex = -1;
  processingStatus: { [key: number]: string } = {};

  constructor(private svc: BatchUploadService) {}

  onFilesSelected(ev: Event) {
    const input = ev.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;

    if (input.files.length > 10) {
      alert('You can only upload a maximum of 10 files at once.');
      input.value = '';
      this.files = [];
      return;
    }

    const selectedFiles = Array.from(input.files);
    const validExtensions = /\.(pptx|ppt|pps|pot|potx|ppsx)$/i;

    for (const file of selectedFiles) {
      if (!validExtensions.test(file.name)) {
        alert(`File "${file.name}" is not a valid PowerPoint file.`);
        input.value = '';
        this.files = [];
        return;
      }
    }

    this.files = selectedFiles;
    this.error = undefined;
  }

  upload() {
    this.error = undefined;
    this.results = [];
    this.uploadProgress = 0;
    this.isProcessing = true;
    this.processingStatus = {};
    this.currentFileIndex = -1;

    if (!this.files.length) {
      this.error = 'No files selected';
      this.isProcessing = false;
      return;
    }

    if (this.files.length > 10) {
      this.error = 'Too many files selected (Max 10)';
      this.isProcessing = false;
      return;
    }

    for (let i = 0; i < this.files.length; i++) {
      this.processingStatus[i] = 'Queued';
    }

    this.svc.batchUpload(this.files).subscribe({
      next: (event: any) => {
        if (event.type === HttpEventType.UploadProgress) {
          this.uploadProgress = Math.round((100 * event.loaded) / (event.total ?? 1));
        } else if (event.type === HttpEventType.Response) {
          const body = event.body || {};

          // backend should return { files: [...] }
          this.results = body.files || [];

          for (let i = 0; i < this.results.length; i++) {
            const result = this.results[i];
            if (result.error) {
              this.processingStatus[i] = `Error: ${result.error}`;
            } else if (result.report) {
              this.processingStatus[i] = '✅ Completed';
            } else {
              this.processingStatus[i] = 'Done';
            }
          }

          this.isProcessing = false;
          this.uploadProgress = 100;
        }
      },
      error: (err: any) => {
        this.isProcessing = false;
        this.error =
          err?.error?.detail ||
          err?.error?.message ||
          err?.message ||
          err?.statusText ||
          'Upload failed';
      },
    });
  }

  downloadFile(suggestedFileName: string) {
    if (!suggestedFileName) return;

    this.error = undefined;

    this.svc.downloadFile(suggestedFileName).subscribe({
      next: (blob) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = suggestedFileName;
        a.click();
        URL.revokeObjectURL(url);
      },
      error: (err: any) => {
        this.error = `Download failed: ${err?.error?.detail || err?.message || 'error'}`;
      },
    });
  }

  downloadAll() {
    this.error = undefined;

    const fileNames = this.results
      .filter((r: any) => r?.suggestedFileName && !r?.error)
      .map((r: any) => r.suggestedFileName);
    console.log('fileNames to download:', fileNames);

    if (!fileNames.length) {
      this.error = 'No downloadable remediated files found';
      return;
    }

    this.svc.downloadAllFiles(fileNames).subscribe({
      next: (blob) => {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'remediated-files.zip';
        a.click();
        URL.revokeObjectURL(url);
      },
      error: (err: any) => {
        this.error = `Batch download failed: ${err?.error?.detail || err?.message || 'error'}`;
      },
    });
  }

  getTotalFixed(): number {
    return this.results.reduce((sum, result) => {
      return sum + (result.report?.summary?.fixed || 0);
    }, 0);
  }

  getVisibleAutoFixed(result: any) {
    return (result?.report?.details?.autoFixedAltText || [])
      .filter((fix: any) => fix.fix !== 'truncatedAltText')
      .sort((a: any, b: any) => a.slideNumber - b.slideNumber);
  }

  hasDownloadableFiles(): boolean {
    return (
      this.results &&
      this.results.length > 0 &&
      this.results.some((r: any) => r.suggestedFileName && !r.error)
    );
  }

  clearAll(fileInput?: HTMLInputElement) {
    this.files = [];
    this.results = [];
    this.error = undefined;
    this.uploadProgress = 0;
    this.isProcessing = false;
    this.processingStatus = {};
    this.currentFileIndex = -1;

    if (fileInput) {
      fileInput.value = '';
    }
  }

  ngOnDestroy() {}
}