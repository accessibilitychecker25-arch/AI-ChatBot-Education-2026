import {
  Component,
  EventEmitter,
  Output,
  Input,
  ViewChild,
  ElementRef,
} from '@angular/core';
import { CommonModule } from '@angular/common';

@Component({
  selector: 'app-file-upload',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './file-upload.component.html',
  styleUrls: ['./file-upload.component.css'],
})
export class FileUploadComponent {
  @Output() submitted = new EventEmitter<{ file: File; title: string }>();
  @Output() submittedMultiple = new EventEmitter<File[]>();
  @Output() cleared = new EventEmitter<void>();
  @Input() hasResults = false;
  @ViewChild('fileInput') fileInput!: ElementRef<HTMLInputElement>;
  selectedFile?: File;
  selectedFiles: File[] = [];

  // Reset file input value to allow reselecting the same file
  resetNativeInputValue() {
    if (this.fileInput?.nativeElement) {
      this.fileInput.nativeElement.value = '';
    }
  }

  // Trigger the file input dialog
  triggerFileDialog() {
    this.fileInput?.nativeElement?.click();
  }

  // Clear selected file
  clearFile() {
    this.selectedFile = undefined;
    this.selectedFiles = [];
    if (this.fileInput?.nativeElement) {
      this.fileInput.nativeElement.value = '';
    }
    this.cleared.emit();
  }

  // Handle file selection
  onFileSelected(event: Event) {
    const input = event.target as HTMLInputElement;
    const files = input.files ? Array.from(input.files) : [];
    const valid = files.filter((f) => this.isValidFile(f));

    this.selectedFiles = valid;
    this.selectedFile = valid[0];
  }

  // Check if the selected file is valid (DOCX OR PPTX)
  isValidFile(file: File): boolean {
    const name = file.name.toLowerCase();
    const type = file.type;

    // 1. Check by Extension (Most reliable)
    if (
      name.endsWith('.docx') ||
      name.endsWith('.pptx') ||
      name.endsWith('.ppt') ||
      name.endsWith('.pps') ||
      name.endsWith('.potx')
    ) {
      return true;
    }

    // 2. Check by MIME Type (Fallback)
    // Word
    if (type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') return true;
    // PowerPoint (Modern)
    if (type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation') return true;
    // PowerPoint (Older)
    if (type === 'application/vnd.ms-powerpoint') return true;

    return false;
  }
  submit() {
    if (!this.isFormValid()) return;

    if (this.selectedFiles.length > 1) {
      this.submittedMultiple.emit(this.selectedFiles.slice());
    } else {
      this.submitted.emit({ file: this.selectedFile!, title: '' });
    }

    this.selectedFile = undefined;
    this.selectedFiles = [];
    this.resetNativeInputValue();
  }

  // Validate if a file is selected
  isFormValid(): boolean {
    return this.selectedFiles && this.selectedFiles.length > 0;
  }
}