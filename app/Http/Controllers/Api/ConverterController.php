<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use App\Services\PdfToExcelConverter;
use Illuminate\Http\Request;
use Illuminate\Http\JsonResponse;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\Validator;
use Illuminate\Support\Str;

class ConverterController extends Controller
{
    protected $converter;

    public function __construct(PdfToExcelConverter $converter)
    {
        $this->converter = $converter;
    }

    /**
     * Upload and convert PDF to Excel
     */
    public function convert(Request $request): JsonResponse
    {
        $validator = Validator::make($request->all(), [
            'pdf_file' => 'required|file|mimes:pdf|max:20480', // 20MB max for bank statements
            'password' => 'nullable|string|max:255', // Optional password for encrypted PDFs
        ]);

        if ($validator->fails()) {
            return response()->json([
                'success' => false,
                'message' => 'Validation failed',
                'errors' => $validator->errors()
            ], 422);
        }

        try {
            $pdfFile = $request->file('pdf_file');
            $password = $request->input('password');
            $originalName = $pdfFile->getClientOriginalName();
            $fileName = Str::random(40) . '.pdf';
            
            // Store the uploaded PDF
            $pdfPath = $pdfFile->storeAs('uploads', $fileName, 'local');
            
            // Convert PDF to Excel with optional password
            $excelPath = $this->converter->convert($pdfPath, $originalName, $password);
            
            // Generate Excel filename for download
            $excelFileName = basename($excelPath);
            
            // Generate download URL
            $downloadUrl = url('/api/converter/download/' . $excelFileName);
            
            return response()->json([
                'success' => true,
                'message' => 'PDF converted to Excel successfully',
                'download_url' => $downloadUrl,
                'original_filename' => $originalName,
                'excel_filename' => $excelFileName,
                'file_type' => $this->detectFileType($originalName),
                'file_path' => $excelPath
            ]);

        } catch (\Exception $e) {
            // Clean up uploaded file if conversion fails
            if (isset($pdfPath)) {
                Storage::disk('local')->delete($pdfPath);
            }
            
                // Check if it's a password-related error
                if (str_contains(strtolower($e->getMessage()), 'password') || 
                    str_contains(strtolower($e->getMessage()), 'encrypted') ||
                    str_contains(strtolower($e->getMessage()), 'locked')) {
                    
                    if (!$password) {
                        return response()->json([
                            'success' => false,
                            'message' => 'This PDF is password protected. Please provide the password.'
                        ], 400);
                    } else {
                        return response()->json([
                            'success' => false,
                            'message' => 'Sorry, this PDF uses encryption that cannot be processed by our current system. Please remove the password from your PDF file and try again.'
                        ], 400);
                    }
                }
            
            return response()->json([
                'success' => false,
                'message' => 'Conversion failed: ' . $e->getMessage()
            ], 500);
        }
    }

    /**
     * Detect file type based on filename
     */
    protected function detectFileType(string $filename): string
    {
        $filenameLower = strtolower($filename);
        
        if (strpos($filenameLower, 'bank') !== false || 
            strpos($filenameLower, 'statement') !== false ||
            strpos($filenameLower, 'account') !== false) {
            return 'bank_statement';
        }
        
        if (strpos($filenameLower, 'invoice') !== false ||
            strpos($filenameLower, 'bill') !== false) {
            return 'invoice';
        }
        
        if (strpos($filenameLower, 'report') !== false ||
            strpos($filenameLower, 'financial') !== false) {
            return 'financial_report';
        }
        
        return 'general';
    }

    /**
     * Download converted Excel file
     */
    public function download(string $file): \Symfony\Component\HttpFoundation\StreamedResponse
    {
        // Try to find the file in date-based folders
        $filePath = $this->findFileInDateFolders($file);
        
        if (!$filePath || !Storage::disk('local')->exists($filePath)) {
            abort(404, 'File not found');
        }

        return Storage::disk('local')->download($filePath, $file);
    }
    
    /**
     * Find file in date-based folders
     */
    protected function findFileInDateFolders(string $fileName): ?string
    {
        try {
            $basePath = 'converted';
            $dateFolders = Storage::disk('local')->directories($basePath);
            
            // Search in date folders (most recent first)
            foreach (array_reverse($dateFolders) as $folder) {
                $filePath = $folder . '/' . $fileName;
                if (Storage::disk('local')->exists($filePath)) {
                    return $filePath;
                }
            }
        } catch (\Exception $e) {
            \Log::warning("Failed to find file: " . $e->getMessage());
        }
        
        return null;
    }

    /**
     * Get conversion status (for future async processing)
     */
    public function status(string $jobId): JsonResponse
    {
        // This can be implemented for async processing
        return response()->json([
            'success' => true,
            'status' => 'completed',
            'message' => 'Conversion completed'
        ]);
    }
}
