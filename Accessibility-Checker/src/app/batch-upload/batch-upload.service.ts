// import { Injectable } from '@angular/core';
// import { HttpClient, HttpEvent } from '@angular/common/http';
// import { Observable } from 'rxjs';
// import { environment } from '../../environments/environment';

// @Injectable({ providedIn: 'root' })
// export class BatchUploadService {
//   constructor(private http: HttpClient) {}

//   // These might show 404s in the console if your Python server 
//   // doesn't have session routes yet, but that is okay for now.
//   createSession(): Observable<{ sessionId: string; expiresInSeconds: number }> {
//     const url = `${environment.apiUrl}/api/session`;
//     return this.http.post<{ sessionId: string; expiresInSeconds: number }>(url, {});
//   }

//   keepAlive(sessionId: string): Observable<any> {
//     const url = `${environment.apiUrl}/api/session`;
//     return this.http.post(url, {}, { headers: { 'X-Session-ID': sessionId } as any });
//   }

//   // --- THIS IS THE FIXED FUNCTION ---
//   batchUpload(files: File[], sessionId?: string): Observable<HttpEvent<any>> {
//     // 1. Point directly to your Python /upload route
//     // Note: ensure environment.apiUrl is 'http://localhost:5000'
//     const url = `${environment.apiUrl}/upload`; 

//     const fd = new FormData();
    
//     // 2. Loop through files and append them with the name 'files'
//     // Python expects 'files', NOT 'files[]' or 'file'
//     files.forEach((f) => {
//       fd.append('files', f);
//     });

//     // We send session ID just in case, though Python might ignore it for now
//     if (sessionId) {
//       fd.append('sessionId', sessionId);
//     }

//     return this.http.post(url, fd, { 
//       observe: 'events', 
//       reportProgress: true 
//     });
//   }

//   getBatchDownload(sessionId: string): Observable<Blob> {
//     // adjusted to match a potential future python endpoint
//     const url = `${environment.apiUrl}/download?sessionId=${encodeURIComponent(sessionId)}`;
//     return this.http.get(url, { responseType: 'blob' });
//   }

//   downloadFile(fileName: string): Observable<Blob> {
//     // Download a single remediated file by name
//     const url = `${environment.apiUrl}/download`;
//     return this.http.post(url, { fileName }, { responseType: 'blob' });
//   }
// }

import { Injectable } from '@angular/core';
import { HttpClient, HttpEvent } from '@angular/common/http';
import { Observable } from 'rxjs';
import { environment } from '../../environments/environment';

@Injectable({ providedIn: 'root' })
export class BatchUploadService {
  constructor(private http: HttpClient) {}

  batchUpload(files: File[]): Observable<HttpEvent<any>> {
    const url = `${environment.apiUrl}${environment.uploadEndpoint}`;
    const fd = new FormData();

    files.forEach((f) => {
      fd.append('files', f);
    });

    return this.http.post(url, fd, {
      observe: 'events',
      reportProgress: true,
    });
  }

  downloadFile(fileName: string): Observable<Blob> {
    const url = `${environment.apiUrl}${environment.downloadEndpoint}`;
    return this.http.post(
      url,
      { fileName },
      { responseType: 'blob' }
    );
  }

  downloadAllFiles(fileNames: string[]): Observable<Blob> {
    const url = `${environment.apiUrl}${environment.downloadEndpoint}`;
    return this.http.post(
      url,
      { files: fileNames },
      { responseType: 'blob' }
    );
  }
}