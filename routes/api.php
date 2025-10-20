<?php

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\Api\ConverterController;

/*
|--------------------------------------------------------------------------
| API Routes
|--------------------------------------------------------------------------
|
| Here is where you can register API routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "api" middleware group. Make something great!
|
*/

Route::middleware('auth:sanctum')->get('/user', function (Request $request) {
    return $request->user();
});

// PDF to Excel Converter Routes
Route::prefix('converter')->group(function () {
    Route::post('/upload', [ConverterController::class, 'convert'])->name('api.convert');
    Route::get('/download/{file}', [ConverterController::class, 'download'])->name('api.download');
    Route::get('/status/{jobId}', [ConverterController::class, 'status'])->name('api.status');
});

// Health check endpoint
Route::get('/health', function () {
    return response()->json([
        'status' => 'ok',
        'timestamp' => now(),
        'service' => 'PDF to Excel Converter API'
    ]);
});
