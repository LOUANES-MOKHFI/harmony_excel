<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelController;
use App\Http\Controllers\RecordController;
use App\Http\Controllers\LoginController;
use App\Http\Controllers\RegisterController;
use App\Http\Controllers\GsmMailController;
use App\Http\Controllers\FaxMobileController;
use App\Http\Controllers\ExcelVicidialcontroller;


/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('', [LoginController::class, 'index'])->name('login');
Route::post('custom-login', [LoginController::class, 'customLogin'])->name('login.custom'); 
Route::get('registration', [RegisterController::class, 'registration'])->name('register-user');
Route::post('custom-registration', [RegisterController::class, 'customRegistration'])->name('register.custom'); 
Route::get('signout', [LoginController::class, 'signOut'])->name('signout');

//Route::post("simple-excel/import", [ExcelController::class,'import'])->name('excel.import');

// Exporter un fichier Excel
//Route::post("simple-excel/export", [ExcelController::class,'export'])->name('excel.export');
Route::get("call1-export", [ExcelController::class,'index'])->name('home');
Route::post("simple-excel/editFile", [ExcelController::class,'editFile'])->name('excel.editFile');


Route::get("call2-export", [ExcelController::class,'call2_export'])->name('call2');
Route::post("call2-export", [ExcelController::class,'editFileCall2'])->name('editFileCall2');


Route::get("records", [RecordController::class,'index'])->name('records');
Route::post("download-records", [RecordController::class,'getAllRecords'])->name('downloadRecords');

Route::get("gsm_mail", [GsmMailController::class,'index'])->name('gsm_mail');
Route::post("download-gsm_mail", [GsmMailController::class,'getAllgsm_mail'])->name('downloadgsm_mail');

Route::get("fax_mobile", [FaxMobileController::class,'index'])->name('fax_mobile');
Route::post("download-fax_mobile", [FaxMobileController::class,'getAllfax_mobile'])->name('downloadfax_mobile');


Route::get("stat_vicidial", [ExcelVicidialcontroller::class,'index'])->name('vicidial_index');
Route::post("vicidial/stat", [ExcelVicidialcontroller::class,'StatExcel'])->name('vicidial_stat');
