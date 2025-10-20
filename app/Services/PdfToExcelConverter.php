<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use Smalot\PdfParser\Parser;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Str;

class PdfToExcelConverter
{
    protected $parser;

    public function __construct()
    {
        $this->parser = new Parser();
    }

    /**
     * Convert PDF file to Excel
     */
    public function convert(string $pdfPath, string $originalName): string
    {
        try {
            // Parse PDF content
            $pdf = $this->parser->parseFile(Storage::disk('local')->path($pdfPath));
            $text = $pdf->getText();
            
            // Extract structured data from PDF text
            $data = $this->extractStructuredData($text);
            
            // Create Excel file
            $excelPath = $this->createExcelFile($data, $originalName);
            
            // Clean up uploaded PDF
            Storage::disk('local')->delete($pdfPath);
            
            return $excelPath;
            
        } catch (\Exception $e) {
            // Clean up uploaded PDF on error
            Storage::disk('local')->delete($pdfPath);
            throw new \Exception('PDF parsing failed: ' . $e->getMessage());
        }
    }

    /**
     * Extract structured data from PDF text
     */
    protected function extractStructuredData(string $text): array
    {
        $lines = array_filter(explode("\n", $text), function($line) {
            return trim($line) !== '';
        });

        $data = [];
        $isBankStatement = $this->detectBankStatement($text);
        
        if ($isBankStatement) {
            $data = $this->extractBankStatementData($lines);
        } else {
            $data = $this->extractGeneralTabularData($lines);
        }

        return $data;
    }

    /**
     * Detect if this is a bank statement
     */
    protected function detectBankStatement(string $text): bool
    {
        $bankKeywords = [
            'bank statement', 'account statement', 'statement of account',
            'transaction history', 'account summary', 'balance',
            'deposit', 'withdrawal', 'transfer', 'payment',
            'debit', 'credit', 'available balance', 'current balance',
            'account number', 'routing number', 'checking', 'savings'
        ];

        $textLower = strtolower($text);
        $keywordCount = 0;

        foreach ($bankKeywords as $keyword) {
            if (strpos($textLower, $keyword) !== false) {
                $keywordCount++;
            }
        }

        return $keywordCount >= 3;
    }

    /**
     * Extract bank statement data with proper formatting
     */
    protected function extractBankStatementData(array $lines): array
    {
        $data = [];
        $headers = ['Date', 'Description', 'Debit', 'Credit', 'Balance'];
        $data[] = $headers;
        
        $transactionLines = [];
        $inTransactionSection = false;
        
        foreach ($lines as $line) {
            $line = trim($line);
            
            // Skip empty lines
            if (empty($line)) {
                continue;
            }
            
            // Detect start of transaction section
            if ($this->isTransactionHeader($line)) {
                $inTransactionSection = true;
                continue;
            }
            
            // Skip headers and summary lines
            if (!$inTransactionSection && !$this->isTransactionLine($line)) {
                continue;
            }
            
            // Extract transaction data
            if ($inTransactionSection && $this->isTransactionLine($line)) {
                $transaction = $this->parseTransactionLine($line);
                if ($transaction) {
                    $transactionLines[] = $transaction;
                }
            }
        }
        
        // Sort transactions by date (if dates are available)
        usort($transactionLines, function($a, $b) {
            if (isset($a['date']) && isset($b['date'])) {
                return strtotime($a['date']) - strtotime($b['date']);
            }
            return 0;
        });
        
        // Add transactions to data
        foreach ($transactionLines as $transaction) {
            $data[] = [
                $transaction['date'] ?? '',
                $transaction['description'] ?? '',
                $transaction['debit'] ?? '',
                $transaction['credit'] ?? '',
                $transaction['balance'] ?? ''
            ];
        }
        
        return $data;
    }

    /**
     * Extract general tabular data
     */
    protected function extractGeneralTabularData(array $lines): array
    {
        $data = [];
        
        foreach ($lines as $line) {
            $line = trim($line);
            
            // Skip empty lines
            if (empty($line)) {
                continue;
            }
            
            // Check if line contains tabular data (numbers, dates, etc.)
            if ($this->isTabularData($line)) {
                $columns = $this->splitIntoColumns($line);
                $data[] = $columns;
            } else {
                // Handle non-tabular data
                $data[] = [$line];
            }
        }

        return $data;
    }

    /**
     * Check if line is a transaction header
     */
    protected function isTransactionHeader(string $line): bool
    {
        $headerPatterns = [
            '/date.*description.*amount/i',
            '/transaction.*date.*description/i',
            '/date.*transaction.*amount/i',
            '/date.*description.*debit.*credit/i'
        ];

        foreach ($headerPatterns as $pattern) {
            if (preg_match($pattern, $line)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Check if line contains transaction data
     */
    protected function isTransactionLine(string $line): bool
    {
        // Look for date patterns
        $datePatterns = [
            '/\d{1,2}\/\d{1,2}\/\d{2,4}/',
            '/\d{4}-\d{2}-\d{2}/',
            '/\d{1,2}-\d{1,2}-\d{2,4}/',
            '/[A-Za-z]{3}\s+\d{1,2},?\s+\d{4}/'
        ];

        $hasDate = false;
        foreach ($datePatterns as $pattern) {
            if (preg_match($pattern, $line)) {
                $hasDate = true;
                break;
            }
        }

        // Look for currency amounts
        $currencyPatterns = [
            '/\$[\d,]+\.?\d*/',
            '/[\d,]+\.\d{2}/',
            '/[\d,]+\.\d{2}\s*[+-]?/'
        ];

        $hasCurrency = false;
        foreach ($currencyPatterns as $pattern) {
            if (preg_match($pattern, $line)) {
                $hasCurrency = true;
                break;
            }
        }

        return $hasDate && $hasCurrency;
    }

    /**
     * Parse transaction line into structured data
     */
    protected function parseTransactionLine(string $line): ?array
    {
        // Extract date
        $date = $this->extractDate($line);
        
        // Extract amounts (look for currency patterns)
        $amounts = $this->extractAmounts($line);
        
        // Extract description (everything that's not date or amount)
        $description = $this->extractDescription($line, $date, $amounts);
        
        // Determine debit/credit
        $debit = '';
        $credit = '';
        $balance = '';
        
        if (!empty($amounts)) {
            // Simple logic: negative amounts are debits, positive are credits
            // This can be enhanced based on specific bank statement formats
            $amount = $amounts[0];
            if (strpos($amount, '-') !== false || strpos($line, 'debit') !== false) {
                $debit = $amount;
            } else {
                $credit = $amount;
            }
            
            // If there are multiple amounts, the last one might be balance
            if (count($amounts) > 1) {
                $balance = end($amounts);
            }
        }

        return [
            'date' => $date,
            'description' => $description,
            'debit' => $debit,
            'credit' => $credit,
            'balance' => $balance
        ];
    }

    /**
     * Extract date from line
     */
    protected function extractDate(string $line): string
    {
        $datePatterns = [
            '/(\d{1,2}\/\d{1,2}\/\d{2,4})/',
            '/(\d{4}-\d{2}-\d{2})/',
            '/(\d{1,2}-\d{1,2}-\d{2,4})/',
            '/([A-Za-z]{3}\s+\d{1,2},?\s+\d{4})/'
        ];

        foreach ($datePatterns as $pattern) {
            if (preg_match($pattern, $line, $matches)) {
                return $matches[1];
            }
        }

        return '';
    }

    /**
     * Extract amounts from line
     */
    protected function extractAmounts(string $line): array
    {
        $amounts = [];
        $currencyPatterns = [
            '/\$[\d,]+\.?\d*/',
            '/[\d,]+\.\d{2}/'
        ];

        foreach ($currencyPatterns as $pattern) {
            if (preg_match_all($pattern, $line, $matches)) {
                $amounts = array_merge($amounts, $matches[0]);
            }
        }

        return array_unique($amounts);
    }

    /**
     * Extract description from line
     */
    protected function extractDescription(string $line, string $date, array $amounts): string
    {
        $description = $line;
        
        // Remove date
        if (!empty($date)) {
            $description = str_replace($date, '', $description);
        }
        
        // Remove amounts
        foreach ($amounts as $amount) {
            $description = str_replace($amount, '', $description);
        }
        
        // Clean up extra spaces and common words
        $description = preg_replace('/\s+/', ' ', trim($description));
        $description = preg_replace('/\b(debit|credit|balance|amount)\b/i', '', $description);
        
        return trim($description);
    }

    /**
     * Check if line contains tabular data
     */
    protected function isTabularData(string $line): bool
    {
        // Look for patterns that suggest tabular data
        $patterns = [
            '/\d+\.\d+/',           // Decimal numbers
            '/\d{1,2}\/\d{1,2}\/\d{4}/', // Dates
            '/\$[\d,]+\.?\d*/',     // Currency
            '/\d{4}-\d{2}-\d{2}/', // ISO dates
        ];

        foreach ($patterns as $pattern) {
            if (preg_match($pattern, $line)) {
                return true;
            }
        }

        // Check for multiple spaces (potential column separators)
        return preg_match('/\s{2,}/', $line);
    }

    /**
     * Split line into columns
     */
    protected function splitIntoColumns(string $line): array
    {
        // Split by multiple spaces or tabs
        $columns = preg_split('/\s{2,}|\t/', $line);
        
        // Clean up columns
        return array_map('trim', array_filter($columns));
    }

    /**
     * Create Excel file from extracted data
     */
    protected function createExcelFile(array $data, string $originalName): string
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        // Set sheet title
        $sheet->setTitle('PDF Data');
        
        // Add data to sheet
        $row = 1;
        foreach ($data as $rowData) {
            $col = 'A';
            foreach ($rowData as $cellData) {
                $sheet->setCellValue($col . $row, $cellData);
                $col++;
            }
            $row++;
        }
        
        // Apply formatting
        $this->applyFormatting($sheet, $row - 1);
        
        // Generate filename
        $fileName = Str::random(40) . '.xlsx';
        $filePath = 'converted/' . $fileName;
        
        // Save Excel file
        $writer = new Xlsx($spreadsheet);
        $writer->save(Storage::disk('local')->path($filePath));
        
        return $filePath;
    }

    /**
     * Apply formatting to Excel sheet
     */
    protected function applyFormatting($sheet, int $maxRow): void
    {
        // Auto-size columns
        foreach (range('A', 'Z') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }
        
        // Apply header formatting if we have multiple rows
        if ($maxRow > 1) {
            $headerRange = 'A1:' . $sheet->getHighestColumn() . '1';
            
            $sheet->getStyle($headerRange)->applyFromArray([
                'font' => [
                    'bold' => true,
                    'color' => ['rgb' => 'FFFFFF']
                ],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['rgb' => '366092']
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical' => Alignment::VERTICAL_CENTER
                ],
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THIN,
                        'color' => ['rgb' => '000000']
                    ]
                ]
            ]);
        }
        
        // Apply borders to all data
        $dataRange = 'A1:' . $sheet->getHighestColumn() . $maxRow;
        $sheet->getStyle($dataRange)->applyFromArray([
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => 'CCCCCC']
                ]
            ]
        ]);
        
        // Apply special formatting for bank statements
        $this->applyBankStatementFormatting($sheet, $maxRow);
    }

    /**
     * Apply special formatting for bank statements
     */
    protected function applyBankStatementFormatting($sheet, int $maxRow): void
    {
        // Check if this looks like a bank statement by examining headers
        $firstRow = $sheet->rangeToArray('A1:' . $sheet->getHighestColumn() . '1')[0];
        $isBankStatement = false;
        
        foreach ($firstRow as $cell) {
            if (in_array(strtolower(trim($cell)), ['date', 'description', 'debit', 'credit', 'balance'])) {
                $isBankStatement = true;
                break;
            }
        }
        
        if (!$isBankStatement) {
            return;
        }
        
        // Format date column (usually column A)
        if ($maxRow > 1) {
            $dateRange = 'A2:A' . $maxRow;
            $sheet->getStyle($dateRange)->getNumberFormat()->setFormatCode('mm/dd/yyyy');
        }
        
        // Format currency columns (look for debit, credit, balance columns)
        $highestColumn = $sheet->getHighestColumn();
        $columnIndex = 1;
        
        for ($col = 'A'; $col <= $highestColumn; $col++) {
            $headerValue = strtolower(trim($sheet->getCell($col . '1')->getValue()));
            
            if (in_array($headerValue, ['debit', 'credit', 'balance']) && $maxRow > 1) {
                $currencyRange = $col . '2:' . $col . $maxRow;
                $sheet->getStyle($currencyRange)->getNumberFormat()->setFormatCode('$#,##0.00');
                
                // Apply conditional formatting for debits (red) and credits (green)
                if ($headerValue === 'debit') {
                    $sheet->getStyle($currencyRange)->getFont()->getColor()->setRGB('CC0000');
                } elseif ($headerValue === 'credit') {
                    $sheet->getStyle($currencyRange)->getFont()->getColor()->setRGB('006600');
                }
            }
            
            $columnIndex++;
        }
        
        // Add alternating row colors for better readability
        for ($row = 2; $row <= $maxRow; $row++) {
            if ($row % 2 === 0) {
                $rowRange = 'A' . $row . ':' . $highestColumn . $row;
                $sheet->getStyle($rowRange)->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()->setRGB('F8F9FA');
            }
        }
        
        // Freeze the header row
        $sheet->freezePane('A2');
    }
}
